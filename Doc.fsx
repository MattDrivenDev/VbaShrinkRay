#I "./packages/fsharp.formatting/lib/net40/"
#I "./packages/razorengine/lib/net45/"
#r "FSharp.CodeFormat.dll"
#r "FSharp.Literate.dll"
#r "RazorEngine.dll"
open FSharp.Literate
open System.IO


let source = __SOURCE_DIRECTORY__
let template = Path.Combine(source, "template.html")

let script = Path.Combine(source, "VbaShrinkRay.fsx")
Literate.ProcessScriptFile(script, template)