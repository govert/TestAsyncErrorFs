module AsyncCallInfo

open ExcelDna.Integration

// Converted from the C# code here: https://github.com/Excel-DNA/ExcelDna/blob/56c86af095da8fbce134a8c30f97ad4ba2b68a60/Source/ExcelDna.Integration/ExcelRtdObserver.cs#L335
// We basically need a deep equality check for all the possible parameter types
// This includes 1D and 2D double and obj arrays
// We'll be putting these objects in a HashSet, so we need to define a hash code as well

// We need to define some helper functions to compare arrays and objects
// The generic constraints mean they can't be defined in the AsyncCallInfo type

let arrayEquals<'T when 'T : equality> (a: 'T[]) (b: 'T[]) =
    if a.Length <> b.Length then false
    else a |> Array.forall2 (=) b

let arrayEquals2D<'T when 'T : equality> (a: 'T[,]) (b: 'T[,]) =
    if a.GetLength(0) <> b.GetLength(0) || a.GetLength(1) <> b.GetLength(1) then false
    else
        let mutable equals = true
        for i in 0 .. a.GetLength(0) - 1 do
            for j in 0 .. a.GetLength(1) - 1 do
                if a.[i, j] <> b.[i, j] then
                    equals <- false
                    // would be nice to return directly here, but F# doesn't allow it easily
        equals

let rec valueEquals (a: obj) (b: obj) =
    let objArrayEquals (a: obj[]) (b: obj[]) =
        if a.Length <> b.Length then false
        else Array.forall2 valueEquals a b

    let objArray2DEquals(a: obj[,]) (b: obj[,]) =
        if a.GetLength(0) <> b.GetLength(0) || a.GetLength(1) <> b.GetLength(1) then false
        else
            let mutable equals = true
            for i in 0 .. a.GetLength(0) - 1 do
                for j in 0 .. a.GetLength(1) - 1 do
                    if not (valueEquals a.[i, j] b.[i, j]) then
                        equals <- false
                        // would be nice to return directly here, but F# doesn't allow it easily
            equals

    if a = b then true
    elif a :? double[] && b :? double[] then arrayEquals (a :?> double[]) (b :?> double[])
    elif a :? double[,] && b :? double[,] then arrayEquals2D (a :?> double[,]) (b :?> double[,])
    elif a :? obj[] && b :? obj[] then objArrayEquals (a :?> obj[]) (b :?> obj[])
    elif a :? obj[,] && b :? obj[,] then objArray2DEquals (a :?> obj[,]) (b :?> obj[,])
    elif a :? byte[] && b :? byte[] then arrayEquals (a :?> byte[]) (b :?> byte[])
    else false

[<Struct>]  // I'm not sure this is important, but it's what the C# code does
[<CustomEquality>]
[<NoComparison>]
type AsyncCallInfo = 
    { 
        FunctionName: string; 
        Parameters: obj; 
        HashCode: int 
    } 
    with
    static member Create(functionName: string, parameters: obj) =
        let info = { FunctionName = functionName; Parameters = parameters; HashCode = 0 }
        { info with HashCode = info.ComputeHashCode() }
        
    member this.ComputeHashCode() =
        let rec computeObjHashCode (obj : obj) =
            match obj with
            | null -> 0
            | :? double | :? string | :? bool | :? System.DateTime | :? int | :? uint | :? int64 | :? uint64 | :? int16 | :? uint16 | :? byte | :? sbyte | :? decimal -> obj.GetHashCode()
            | :? ExcelReference | :? ExcelError | :? ExcelEmpty | :? ExcelMissing -> obj.GetHashCode()
            | _ when obj.GetType().IsEnum -> obj.GetHashCode()
            | :? (double[]) as doubles -> 
                Array.fold (fun acc x -> acc * 23 + x.GetHashCode()) 17 doubles
            | :? (double[,]) as doubles2 -> 
                let mutable hash = 17
                for i in 0 .. doubles2.GetLength(0) - 1 do
                    for j in 0 .. doubles2.GetLength(1) - 1 do
                        hash <- hash * 23 + doubles2.[i, j].GetHashCode()
                hash
            | :? (obj[]) as objects -> 
                Array.fold (fun acc x -> acc * 23 + (if x <> null then computeObjHashCode x else 0)) 17 objects
            | :? (obj[,]) as objects2 -> 
                let mutable hash = 17
                for i in 0 .. objects2.GetLength(0) - 1 do
                    for j in 0 .. objects2.GetLength(1) - 1 do
                        let x = objects2.[i, j]
                        hash <- hash * 23 + (if x <> null then computeObjHashCode x else 0) 
                hash
            | :? (byte[]) as bytes -> 
                Array.fold (fun acc x -> acc * 23 + int x) 17 bytes
            | _ -> raise (System.ArgumentException("Invalid type used for async parameter(s)", "parameters"))
        
        (17 * 23 + (if this.FunctionName <> null then this.FunctionName.GetHashCode() else 0)) * 23 + computeObjHashCode this.Parameters

    member this.Equals(other: AsyncCallInfo) =
        this.HashCode = other.HashCode && this.FunctionName = other.FunctionName && valueEquals this.Parameters other.Parameters

    override this.GetHashCode() = this.HashCode

    // Add equality members override 
    override this.Equals(other: obj) =
        match other with
        | :? AsyncCallInfo as other -> this.Equals(other)
        | _ -> false
