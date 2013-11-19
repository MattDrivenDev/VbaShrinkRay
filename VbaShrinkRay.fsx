open System
open System.IO
open System.Text.RegularExpressions

type RawLine = RawLine of string list

type ClassifiedLine = 
    | DeclarationLine of string
    | BodyLine of string list
    
type Definition = Definition of string

type LineOfCode = LineOfCode of string list

type CodeBlock = { 
    SourceFile : string
    Definition : Definition option
    Lines      : LineOfCode list }

let fqn codeblock = 
    match codeblock.Definition with
    | Some (Definition def) -> sprintf "%s.%s" codeblock.SourceFile def
    | None -> sprintf "%s.Global" codeblock.SourceFile
    
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
    
let classifyCode (RawLine words) =
    match words with
    | [] -> failwith "Cannot classify an empty line."
    | _ :: "sub" :: name :: _ -> DeclarationLine name
    | "sub" :: name :: _ -> DeclarationLine name
    | _ :: "function" :: name :: _ -> DeclarationLine name
    | "function" :: name :: _ -> DeclarationLine name
    | _ :: "table" :: name :: _ -> DeclarationLine name
    | "table" :: name :: _ -> DeclarationLine name
    | _ :: "module" :: name :: _ -> DeclarationLine name
    | "module" :: name :: _ -> DeclarationLine name
    | _ -> BodyLine words

let codeBlock filename = 
    function
    | BodyLine words -> {SourceFile=filename;Definition=None;Lines=[LineOfCode words]}
    | DeclarationLine name -> {SourceFile=filename;Definition=Some(Definition name);Lines=[]}

let addLineOfCode codeblock line = 
    {codeblock with Lines=codeblock.Lines @ [line]}

let intoCodeBlocks filename (codeDom:CodeBlock list) line =
    match codeDom with
    | [] -> [codeBlock filename line]
    | head :: tail -> 
        match line with
        | BodyLine words -> addLineOfCode head (LineOfCode words) :: tail
        | DeclarationLine name -> codeBlock filename line :: codeDom

let parse filename =     
    let parseAllLines =
        Seq.map lowerCase
        >> removeCodeGen
        >> Seq.map split
        >> Seq.filter notEmpty
        >> Seq.map RawLine
        >> Seq.map classifyCode
        >> Seq.fold (intoCodeBlocks filename) []    
    filename |> (File.ReadAllLines >> parseAllLines)