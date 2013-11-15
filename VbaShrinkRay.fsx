(**
# VbaShinkRay

Shinking all of your vba's! `>:D`
*)

open System
open System.IO
open System.Text.RegularExpressions
Directory.SetCurrentDirectory(__SOURCE_DIRECTORY__)

type rawLineOfCode = RawLine of string list

type classifiedLineOfCode = 
    | DeclarationLine of string
    | BodyLine of string list
    
type blockDefinition = BlockDefinition of string

type blockLineOfCode = BlockLineOfCode of string list

type codeBlock = { 
    Definition: blockDefinition option
    Lines: blockLineOfCode list }

type parser = string seq -> codeBlock list

let lowerCase (s:string) = s.ToLower()

let notEmpty (xs:seq<string>) = xs |> Seq.isEmpty |> not

let split s = 
    let pattern = @"\w+"
    Regex.Matches(s, pattern)
    |> Seq.cast<Match>
    |> Seq.map (fun m -> m.ToString())
    |> List.ofSeq

let rawLineOfCode = RawLine

let classifiedLineOfCode (RawLine words) =
    match words with
    | [] -> failwith "Cannot classify an empty line."
    | _ :: "sub" :: name :: _ -> DeclarationLine name
    | _ :: "function" :: name :: _ -> DeclarationLine name
    | _ -> BodyLine words

let codeBlocks (codeDom:codeBlock list) line =     
    match codeDom with
    | [] -> 
        match line with
        | BodyLine words -> [{ Definition = None; Lines = [BlockLineOfCode words] }]
        | DeclarationLine name -> [{ Definition = Some(BlockDefinition name); Lines = [] }]
    | head :: tail -> 
        match line with
        | BodyLine words -> { head with Lines = head.Lines @ [BlockLineOfCode words] } :: tail
        | DeclarationLine name -> { Definition = Some(BlockDefinition name); Lines = [] } :: codeDom

let parse : parser =     
    Seq.map lowerCase
    >> Seq.map split
    >> Seq.filter notEmpty
    >> Seq.map rawLineOfCode
    >> Seq.map classifiedLineOfCode
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