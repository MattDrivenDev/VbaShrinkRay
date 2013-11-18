(**
# VbaShinkRay

Shinking all of your vba's! `>:D`
*)

open System
open System.IO
open System.Text.RegularExpressions

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

let removeCodeGen xs =    
    Seq.fold (fun lines (l:string) ->
        if l = "codebehindform" then [||]
        else Array.append lines [|l|]
    ) [||] xs

let notEmpty (xs:seq<string>) = xs |> Seq.isEmpty |> not

let split s = 
    let pattern = @"\w+"
    Regex.Matches(s, pattern)
    |> Seq.cast<Match>
    |> Seq.map (fun m -> m.ToString())
    |> List.ofSeq
    
let classifiedLineOfCode (RawLine words) =
    match words with
    | [] -> failwith "Cannot classify an empty line."
    | _ :: "sub" :: name :: _ -> DeclarationLine name
    | "sub" :: name :: _ -> DeclarationLine name
    | _ :: "function" :: name :: _ -> DeclarationLine name
    | "function" :: name :: _ -> DeclarationLine name
    | _ :: "table" :: name :: _ -> DeclarationLine name
    | "table" :: name :: _ -> DeclarationLine name
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
    >> removeCodeGen
    >> Seq.map split
    >> Seq.filter notEmpty
    >> Seq.map RawLine
    >> Seq.map classifiedLineOfCode
    >> Seq.fold codeBlocks []