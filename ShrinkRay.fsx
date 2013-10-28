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

1. Total number of lines of code
2. Total number of procedures
3. Total number of dependencies upon a procedure, and their locations

There are not a great number of requirements - and I'm going to be taking a 
fairly naive approach to meeting them; but it should be **good enough** to be
useful.

I am going to be working with a fairly standard folder structure full of text
files that contain VBA source code, as generated with [MSAccess SVN][svn]

    ~/src/
       |> Forms/
       |> General/
       |> Modules/
       |> Queries/
       |> Reports/
       |> Scripts/
       |> Tables/

 [svn]: http://accesssvn.codeplex.com/
*)

(*** hide ***)
open System
open System.IO
open System.Text.RegularExpressions
Directory.SetCurrentDirectory(__SOURCE_DIRECTORY__)

(**
First up, a few types that will help define the domain of VBA source code:
*)

/// Defines a raw and unprocessed line of VBA code.
type RawLine = RawLine of string list

/// Defines a line of VBA code during processing, which can be either a line
/// of declaration code (sub, function, module etc.) - or a line of body code.
type ClassifiedLine = 
    | DeclarationLine of string
    | BodyLine of string list

/// Defines a line of VBA code once processing has completed, that represents a
/// definition (sub, function, module etc.) - essentially a codeblock header.
type Definition = Definition of string

/// Defines a line of VBA code once processing has completed, that represents a
/// normal line of code within a block of code.
type LineOfCode = LineOfCode of string list

/// Defines a block of code (essentially a header + body).
type CodeBlock = 
    { /// The 'header' for the codeblock, a sub/function/module definition.
      Definition: Definition
      /// The 'body' for the codeblock, statements comments etc.
      Lines: LineOfCode list }

(**

*)

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