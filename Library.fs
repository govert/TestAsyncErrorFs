module TestAsyncError

open System
open System.Collections.Generic
open System.Threading
open ExcelDna.Integration
open AsyncCallInfo

// This is an example Observable function with an extra flag to indicate whether to fail while running
type ExcelObservableClock(input: obj, fail: bool) =
    // Declare the mutable field for the observer
    let mutable _observer: IExcelObserver = null

    let timerCallback _ =
        if fail then
            // Assuming '_observer' is defined elsewhere in your type
            if _observer <> null then
                _observer.OnError(Exception(sprintf "[%A] Error at %A" input (DateTime.Now.ToString("HH:mm:ss.fff"))))
        else
            if _observer <> null then
                _observer.OnNext(sprintf "[%A] %A" input (DateTime.Now.ToString("HH:mm:ss.fff")))

    let timer = new Timer(timerCallback, null, 0, 1000)


    interface IExcelObservable with
        member this.Subscribe(observer: IExcelObserver) =
            _observer <- observer
            observer.OnNext(sprintf "[%A] %A (Subscribe)" input (DateTime.Now.ToString("HH:mm:ss.fff")))
            { new IDisposable with member this.Dispose() = () }

// This simulates the error fallback UDF which might be defined in a different add-in
let AsyncErrorFallback (input: obj) =
    let functionName = "AsyncErrorFallback"
    let args = [| input |]
    let asyncFuncImpl = 
        fun () -> 
            Thread.Sleep(1000)
            sprintf "[%A] Error result at %A" input (DateTime.Now) :> obj
    let asyncFunc = new ExcelFunc(asyncFuncImpl)
    ExcelAsyncUtil.Run(functionName, args, asyncFunc)



// ################   Observe with Error Fallback start  ################

// This set keeps track of calls that are in an error state, and should direct to the error fallback
let _errorCalls = HashSet<AsyncCallInfo>()

// This is an example Observable function that does async error handling
let RunClock (input: obj) (fail: bool) =
    let functionName = "RunClock"
    let args = [| input; fail :> obj |]
    let callInfo = AsyncCallInfo.Create(functionName, args) // This will be the key in our _errorCalls set

    // We define a local function to call the error fallback, 
    // we also check whether the error fallback has returned a result 
    // (typically it will in the next call) and then remove the call info from the error tracking set, to reset everything
    let callErrorFallback = fun () ->
        // Error handling mode
        let errorFallbackResult = XlCall.Excel(XlCall.xlUDF, "AsyncErrorFallback", input)
        if not (errorFallbackResult.Equals(ExcelError.ExcelErrorNA)) then
            _errorCalls.Remove(callInfo) |> ignore
        errorFallbackResult

    // Our function check whether we are in error handling mode or not
    if not (_errorCalls.Contains(callInfo)) then
        // We are not in error handling mode, so call the real observable
        let result = ExcelAsyncUtil.Observe(functionName, args, (fun () -> ExcelObservableClock(input, fail)))

        // We now check whether the Observable set an error result or not
        if not (result.Equals(ExcelError.ExcelErrorValue)) then
            // Everything is fine, no error from the observable so return the result
            result
        else
            // We have an error from the observable, so we need to go into error handling mode
            _errorCalls.Add(callInfo) |> ignore
            // Call the error fallback and return the result from that call (typically #N/A the first time)
            callErrorFallback()
    else
        // Call the error fallback and return the result
        callErrorFallback()
