(**
# VBA Shrink-Ray!

Being able to easily and regularly visualize how much **legacy code** exists
in your application's codebase - as [Dan Milstein][dm] describes in his article
[Screw you Joel Spolsky, We're Rewriting It From Scratch!][rw]; is a key tool
in measuring success.

This is a script that helps treat a codebase like data - for use in doing a
**re-write**; and targets `VBA` code in a reasonably naive way. But should 
still remain valuable.
 
 [dm]: https://twitter.com/danmil
 [rw]: http://onstartups.com/tabid/3339/bid/97052/Screw-You-Joel-Spolsky-We-re-Rewriting-It-From-Scratch.aspx
*)

(** 
## Requirements

I'm a big fan of requirements. In fact, more than once I've stated that 
**having requirements, is a business requirement**.
*)

(*** hide ***)
open System
open System.IO
open System.Text.RegularExpressions
Directory.SetCurrentDirectory(__SOURCE_DIRECTORY__)

type RawLine = RawLine of string list

type ClassifiedLine = 
    | DeclarationLine of string
    | BodyLine of string list

type Definition = Definition of string

type LineOfCode = LineOfCode of string list

type CodeBlock = {    
    Definition: Definition
    Lines: LineOfCode list
}

let lowerCase (s:string) = s.ToLower()

let split s = 
    let pattern = @"\w+"
    Regex.Matches(s, pattern)
    |> Seq.cast<Match>
    |> Seq.map (fun m -> m.ToString())
    |> List.ofSeq
    
let notEmpty (xs:seq<string>) = xs |> Seq.isEmpty |> not

let classify (RawLine words) =
    match words with
    | [] -> failwith "Cannot classify an empty line."
    | "private" :: "sub" :: name :: _ -> DeclarationLine name
    | "public" :: "sub" :: name :: _ -> DeclarationLine name
    | "private" :: "function" :: name :: _ -> DeclarationLine name
    | "public" :: "function" :: name :: _ -> DeclarationLine name
    | _ -> BodyLine words

let codeBlocks (state:CodeBlock list) line =     
    match state with

    // Let's hope our 'line' is a 'DeclarationLine'
    | [] -> 
        match line with
        | BodyLine _ -> 
            failwith "Cannot create a new code block without a declaration line"
        | DeclarationLine name -> 
            [{ Definition = Definition name; Lines = [] }]

    // The 'head' of our list is the 'current' code block
    | head :: tail -> 
        match line with
        | BodyLine words ->
            { head with Lines = head.Lines @ [LineOfCode words] } :: tail
        | DeclarationLine name -> // The last code block is done!
            { Definition = Definition name; Lines = [] } :: state

let parse : (seq<string>->CodeBlock list) =     
    Seq.map lowerCase
    >> Seq.map split
    >> Seq.filter notEmpty
    >> Seq.map RawLine
    >> Seq.map classify
    >> Seq.fold codeBlocks []

[ "Private Sub ZeroArgumentSubProcedureDoesNothing()"
  "    'intentionally does nothing"
  "End Sub" ]
|> parse
|> Seq.iter (printfn "%A")

[ "" ]
|> parse
|> Seq.iter (printfn "%A")

[ "Dim i = 0" ]
|> parse
|> Seq.iter (printfn "%A")