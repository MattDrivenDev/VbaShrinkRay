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
    | Some (Definition def) -> Some(sprintf "%s.%s" codeblock.SourceFile def)
    | None -> None
    
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

let hasDependency (Definition def) dependant = 
    let exists = 
        dependant.Lines
        |> Seq.map (fun (LineOfCode line) -> line |> Seq.exists ((=)def))
        |> Seq.exists ((=)true)
    if exists then Some dependant else None

let getDefinitions codeblocks = 
    codeblocks 
    |> List.choose (fun block -> 
        match block.Definition with
        | Some def -> Some block
        | None -> None)
    |> Seq.distinct 
    |> Seq.sort 
    |> List.ofSeq

let getDependents codeblocks =
    let findUsages dependency = 
        match dependency.Definition with
        | Some def -> 
            let dependants = codeblocks |> Seq.choose (hasDependency def)            
            dependency, (if Seq.length dependants > 0 then Some dependants else None)
        | None -> 
            dependency, None 
    getDefinitions codeblocks 
    |> Seq.map findUsages
    |> Map.ofSeq

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

let parseFiles filenames = 
    filenames
    |> Array.ofSeq 
    |> Array.concat
    |> Array.Parallel.map parse
    |> Array.fold (@) []


// ----------------------------------------------------------------------------
// EXAMPLE
// ----------------------------------------------------------------------------

let sourceDirs = seq { 
    yield Directory.GetFiles @"access_db_source\forms\"
    yield Directory.GetFiles @"access_db_source\general\"
    yield Directory.GetFiles @"access_db_source\modules\"
    yield Directory.GetFiles @"access_db_source\queries\"
    yield Directory.GetFiles @"access_db_source\reports\"
    yield Directory.GetFiles @"access_db_source\scripts\"
    yield Directory.GetFiles @"access_db_source\tables\" }

sourceDirs 
|> parseFiles 
|> getDependents 
|> Map.iter (fun dependency maybeDependants ->
    let dependencyFqn = fqn dependency
    match dependencyFqn, maybeDependants with
    | Some fqn, Some dependants ->
        let count = dependants |> Seq.length
        printfn "'%s' has %i dependants, which are:" fqn count
    | Some fqn, None -> 
        printfn "'%s' has 0 dependants." fqn
    | None, _ -> 
        printfn "Incorrectly reported dependency in '%s'" dependency.SourceFile)