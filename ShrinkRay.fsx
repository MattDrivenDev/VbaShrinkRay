(**
VBA Shrink-Ray!
===============

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

(*** hide ***)
open System
open System.IO
open System.Text.RegularExpressions

(**
VBA
---

First things first - we need some types to represent the domain of `VBA` code.

We're not going to go too mad here (but feel free to get involved, and extend
what there is!) - I think we can get the most value from just a few simple
constructs.
*)

/// Represent bindings of args, parameters, variables as a name
/// and optional type information.
type allocation = Allocation of string * string option

/// Record of information about a procedure call.
type procedureInfo = 
    { Name: string
      Args: allocation list option
      LineCount: int }

/// Available types of procedure calls.
type procedure =
    | Sub of procedureInfo
    | Function of procedureInfo

(**
With just these few types we can probably generate the following information
pretty easily:

* Total number of `Sub` and `Function` procedures.
* Total number of calls to each `Sub` and `Function` procedures.
* Details of the arguments required by both `Sub` and `Function` procedures.
*)

(*** hide ***)
let bind f a = 
    match a with 
    | Some b -> Some (f b)
    | None -> None

(*** hide ***)
let (>>=) a f = bind f a

(*** hide ***)
let normalizeLineEndings s =
    Regex.Replace(s, @"\r\n|\n\r|\n|\r", "\r\n")

(*** hide ***)
let split (splitter:string) (s:string) = 
    s.Split(separator=[|splitter|], 
            options=StringSplitOptions.RemoveEmptyEntries)
    |> Seq.ofArray

(*** hide ***)
let trim (s:string) = s.Trim()

(*** hide ***)
let trimAll xs = Seq.map trim xs

(*** hide ***)
let removeEmpty xs =
    Seq.filter (String.IsNullOrWhiteSpace >> not) xs

(*** hide ***)
let isAny xs = 
    if Seq.isEmpty xs then None else Some xs

(*** hide ***)
let tokenizeLines = 
    normalizeLineEndings 
    >> split "\r\n"
    >> trimAll
    >> removeEmpty
    >> isAny

(*** hide ***)
let tokenizeWords = 
    split " " >> trimAll

(*** hide ***)
let tokenize vb =
    tokenizeLines vb >>= (Seq.map tokenizeWords)

/// Parses vba code represented as a string and returns a collection
/// of all the discovered procedures.
let getProcedures vb = 
    option<procedure seq>.None

(**
Let's have a play:
*)

let test (results:seq<seq<string>> option) = 
    match results with
    | Some lines ->
        lines 
        |> Seq.iteri (fun ln line ->
            line
            |> Seq.iteri (fun wn word ->
                printfn "Line %i, Word %i: %s" ln wn word
            )
        )
    | None -> printfn "nothing to see here!"

tokenize """
    Sub ZeroArgumentSubProcedure()
        'does nothing
    End Sub
""" |> test

tokenize "" |> test