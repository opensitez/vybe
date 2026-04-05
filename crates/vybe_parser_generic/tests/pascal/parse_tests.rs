/// Pascal parsing tests — every source from the compiler tests run through
/// the generic parser engine. These verify the generic parser is on par
/// with the hand-written vybe_parser_pascal.

use vybe_parser_generic::grammar::*;
use vybe_parser_generic::lexer::tokenize;
use vybe_parser_generic::parser::parse;
use vybe_parser_generic::*;

pub(crate) fn pascal_grammar_pub() -> GrammarDef { pascal_grammar() }

fn pascal_grammar() -> GrammarDef {
    GrammarDef {
        language: LanguageSpec {
            name: "pascal".into(),
            case_sensitive: false,
            statement_terminator: Terminator::Char(';'),
            indentation_based: false,
            expression_language: false,
        },
        lexer: LexerSpec {
            comment_line: vec!["//".into()],
            comment_block: vec![("{".into(), "}".into()), ("(*".into(), "*)".into())],
            string_delimiters: vec!["'".into()],
            string_escape: Some("''".into()),
            triple_string: Vec::new(),
            string_prefixes: Vec::new(),
            interpolation: None,
            template_string: None,
            char_prefix: Some("#".into()),
            hex_prefix: Some("$".into()),
            keywords: vec![
                "program".into(),"unit".into(),"uses".into(),"begin".into(),"end".into(),
                "var".into(),"const".into(),"type".into(),
                "procedure".into(),"function".into(),"constructor".into(),"destructor".into(),"forward".into(),
                "if".into(),"then".into(),"else".into(),
                "for".into(),"to".into(),"downto".into(),"do".into(),"in".into(),
                "while".into(),"repeat".into(),"until".into(),
                "case".into(),"of".into(),"otherwise".into(),
                "class".into(),"record".into(),"interface".into(),"inherited".into(),
                "override".into(),"virtual".into(),"abstract".into(),
                "try".into(),"except".into(),"finally".into(),"raise".into(),"on".into(),
                "and".into(),"or".into(),"not".into(),"xor".into(),"div".into(),"mod".into(),"shl".into(),"shr".into(),
                "nil".into(),"true".into(),"false".into(),
                "exit".into(),"break".into(),"continue".into(),"halt".into(),
                "with".into(),"is".into(),"as".into(),
                "result".into(),"self".into(),
                "public".into(),"private".into(),"protected".into(),"published".into(),
                "array".into(),"set".into(),"file".into(),
                "string".into(),"integer".into(),"real".into(),"boolean".into(),"char".into(),
                "byte".into(),"word".into(),"longint".into(),"shortint".into(),"cardinal".into(),"int64".into(),
                "single".into(),"double".into(),"extended".into(),"pointer".into(),
                "new".into(),"dispose".into(),
            ],
            operators: vec![
                ":=".into(),"+=".into(),"-=".into(),"*=".into(),"/=".into(),
                "<>".into(),"<=".into(),">=".into(),"..".into(),
                "+".into(),"-".into(),"*".into(),"/".into(),
                "=".into(),"<".into(),">".into(),
                "(".into(),")".into(),"[".into(),"]".into(),
                ".".into(),",".into(),";".into(),":".into(),"^".into(),"@".into(),
            ],
        },
        operators: OperatorTable {
            prefix: vec!["not".into(), "-".into(), "@".into()],
            postfix: Vec::new(),
            infix: vec![
                InfixLevel { precedence: 1, ops: vec!["or".into(), "xor".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 2, ops: vec!["and".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 3, ops: vec!["=".into(),"<>".into(),"<".into(),">".into(),"<=".into(),">=".into(),"in".into(),"is".into(),"as".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 4, ops: vec!["+".into(), "-".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 5, ops: vec!["*".into(), "/".into(), "div".into(), "mod".into(), "shl".into(), "shr".into()], assoc: Assoc::Left },
            ],
        },
        blocks: BlockSpec { open: "begin".into(), close: "end".into(), prefix: None, close_with_kind: false },
        types: TypeSpec { position: TypePosition::After, separator: Some(":".into()), return_separator: None },
        statements: Vec::new(),
        declarations: Vec::new(),
        expressions: ExpressionSpec {
            member_access: Some(".".into()),
            optional_chain: None,
            index_open: Some("[".into()),
            index_close: Some("]".into()),
            call_open: Some("(".into()),
            call_close: Some(")".into()),
            deref: Some("^".into()),
            primary_forms: Vec::new(),
        },
        params: ParamSpec {
            open: "(".into(), close: ")".into(), separator: ";".into(),
            name_type_sep: Some(":".into()), type_position: TypePosition::After,
            default_value: Some("=".into()),
            rest_prefix: None, kwargs_prefix: None,
            multi_name: true, multi_name_sep: Some(",".into()),
            pass_by: [("var".into(), "ref".into()), ("const".into(), "const".into())].into_iter().collect(),
        },
        assignment: AssignmentSpec {
            operator: Some(":=".into()),
            compound: [("+=".into(),"Add".into()),("-=".into(),"Sub".into()),("*=".into(),"Mul".into()),("/=".into(),"Div".into())].into_iter().collect(),
            walrus: None,
        },
        program: ProgramSpec { header: None, uses: None, body: None },
    }
}

fn parse_ok(src: &str) -> Module {
    let g = pascal_grammar();
    let tokens = tokenize(src, &g.lexer, &g.language.statement_terminator, false, false);
    parse(&tokens, &g).unwrap_or_else(|e| panic!("parse failed: {}", e))
}

fn parse_succeeds(src: &str) {
    let _ = parse_ok(src);
}

// ═══════════════════════════════════════════════════════════
// LITERALS
// ═══════════════════════════════════════════════════════════

#[test] fn lit_integer()       { parse_succeeds("program T; begin WriteLn(42); end."); }
#[test] fn lit_negative()      { parse_succeeds("program T; begin WriteLn(-7); end."); }
#[test] fn lit_real()          { parse_succeeds("program T; begin WriteLn(3.14); end."); }
#[test] fn lit_string()        { parse_succeeds("program T; begin WriteLn('hello'); end."); }
#[test] fn lit_empty_string()  { parse_succeeds("program T; begin WriteLn(''); end."); }
#[test] fn lit_bool()          { parse_succeeds("program T; begin WriteLn(true); WriteLn(false); end."); }
#[test] fn lit_nil()           { parse_succeeds("program T; begin WriteLn(nil); end."); }

// ═══════════════════════════════════════════════════════════
// VARIABLES & ASSIGNMENT
// ═══════════════════════════════════════════════════════════

#[test] fn var_integer()   { parse_succeeds("program T; var x: Integer; begin x := 10; WriteLn(x); end."); }
#[test] fn var_string()    { parse_succeeds("program T; var s: String; begin s := 'world'; WriteLn(s); end."); }
#[test] fn var_multiple()  { parse_succeeds("program T; var a, b: Integer; begin a := 10; b := 20; WriteLn(a + b); end."); }
#[test] fn var_reassign()  { parse_succeeds("program T; var x: Integer; begin x := 1; x := 2; x := 3; WriteLn(x); end."); }

#[test] fn const_decl()    { parse_succeeds("program T; const N = 42; begin WriteLn(N); end."); }
#[test] fn const_typed()   { parse_succeeds("program T; const N: Integer = 42; begin WriteLn(N); end."); }
#[test] fn const_string()  { parse_succeeds("program T; const S = 'hello'; begin WriteLn(S); end."); }

// ═══════════════════════════════════════════════════════════
// OPERATORS
// ═══════════════════════════════════════════════════════════

#[test] fn op_add()        { parse_succeeds("program T; begin WriteLn(3 + 4); end."); }
#[test] fn op_sub()        { parse_succeeds("program T; begin WriteLn(10 - 3); end."); }
#[test] fn op_mul()        { parse_succeeds("program T; begin WriteLn(6 * 7); end."); }
#[test] fn op_div()        { parse_succeeds("program T; begin WriteLn(10 / 4); end."); }
#[test] fn op_idiv()       { parse_succeeds("program T; begin WriteLn(10 div 3); end."); }
#[test] fn op_mod()        { parse_succeeds("program T; begin WriteLn(10 mod 3); end."); }
#[test] fn op_precedence() { parse_succeeds("program T; begin WriteLn(2 + 3 * 4); end."); }
#[test] fn op_parens()     { parse_succeeds("program T; begin WriteLn((2 + 3) * 4); end."); }
#[test] fn op_and()        { parse_succeeds("program T; begin if true and false then WriteLn('y'); end."); }
#[test] fn op_or()         { parse_succeeds("program T; begin if true or false then WriteLn('y'); end."); }
#[test] fn op_not()        { parse_succeeds("program T; begin if not false then WriteLn('y'); end."); }
#[test] fn op_compare()    { parse_succeeds("program T; begin if 5 = 5 then WriteLn('y'); end."); }
#[test] fn op_ne()         { parse_succeeds("program T; begin if 5 <> 6 then WriteLn('y'); end."); }
#[test] fn op_lt_gt()      { parse_succeeds("program T; begin if 3 < 5 then WriteLn('y'); if 5 > 3 then WriteLn('y'); end."); }
#[test] fn op_le_ge()      { parse_succeeds("program T; begin if 3 <= 3 then WriteLn('y'); if 5 >= 5 then WriteLn('y'); end."); }

// ═══════════════════════════════════════════════════════════
// COMPOUND ASSIGNMENT
// ═══════════════════════════════════════════════════════════

#[test] fn compound_add()  { parse_succeeds("program T; var x: Integer; begin x := 10; x += 5; end."); }
#[test] fn compound_sub()  { parse_succeeds("program T; var x: Integer; begin x := 10; x -= 3; end."); }
#[test] fn compound_mul()  { parse_succeeds("program T; var x: Integer; begin x := 5; x *= 3; end."); }
#[test] fn compound_div()  { parse_succeeds("program T; var x: Real; begin x := 10.0; x /= 4.0; end."); }

// ═══════════════════════════════════════════════════════════
// CONTROL FLOW
// ═══════════════════════════════════════════════════════════

#[test] fn if_then()       { parse_succeeds("program T; begin if true then WriteLn('y'); end."); }
#[test] fn if_else()       { parse_succeeds("program T; begin if false then WriteLn('y') else WriteLn('n'); end."); }
#[test] fn if_nested()     { parse_succeeds("program T; var x: Integer; begin x := 10; if x > 5 then if x > 8 then WriteLn('big') else WriteLn('med'); end."); }
#[test] fn if_block()      { parse_succeeds("program T; begin if true then begin WriteLn('a'); WriteLn('b'); end; end."); }

#[test] fn for_to()        { parse_succeeds("program T; var i: Integer; begin for i := 1 to 5 do WriteLn(i); end."); }
#[test] fn for_downto()    { parse_succeeds("program T; var i: Integer; begin for i := 5 downto 1 do WriteLn(i); end."); }
#[test] fn for_block()     { parse_succeeds("program T; var i, s: Integer; begin s := 0; for i := 1 to 5 do begin s := s + i; end; WriteLn(s); end."); }
#[test] fn for_in()        { parse_succeeds("program T; var x: Integer; begin for x in [10, 20, 30] do WriteLn(x); end."); }

#[test] fn while_basic()   { parse_succeeds("program T; var i: Integer; begin i := 0; while i < 3 do begin WriteLn(i); i := i + 1; end; end."); }
#[test] fn repeat_until()  { parse_succeeds("program T; var i: Integer; begin i := 1; repeat WriteLn(i); i := i + 1; until i > 3; end."); }

#[test] fn case_basic()    { parse_succeeds("program T; var x: Integer; begin x := 2; case x of 1: WriteLn('one'); 2: WriteLn('two'); 3: WriteLn('three'); end; end."); }
#[test] fn case_else()     { parse_succeeds("program T; var x: Integer; begin x := 5; case x of 1: WriteLn('one'); else WriteLn('other'); end; end."); }

#[test] fn break_stmt()    { parse_succeeds("program T; var i: Integer; begin for i := 1 to 10 do begin if i > 3 then break; WriteLn(i); end; end."); }
#[test] fn continue_stmt() { parse_succeeds("program T; var i: Integer; begin for i := 1 to 5 do begin if i = 3 then continue; WriteLn(i); end; end."); }

// ═══════════════════════════════════════════════════════════
// FUNCTIONS & PROCEDURES
// ═══════════════════════════════════════════════════════════

#[test] fn proc_basic() {
    parse_succeeds("program T; procedure Greet(name: String); begin WriteLn('Hello ' + name); end; begin Greet('World'); end.");
}

#[test] fn func_basic() {
    parse_succeeds(r#"program T; function Add(a, b: Integer): Integer; begin Result := a + b; end; begin WriteLn(Add(3, 4)); end."#);
}

#[test] fn func_recursive() {
    parse_succeeds(r#"program T;
function Fact(n: Integer): Integer;
begin if n <= 1 then Result := 1 else Result := n * Fact(n - 1); end;
begin WriteLn(Fact(5)); end."#);
}

#[test] fn func_nested() {
    parse_succeeds(r#"program T;
function Outer(x: Integer): Integer;
  function Inner(y: Integer): Integer; begin Result := y * 2; end;
begin Result := Inner(x) + 1; end;
begin WriteLn(Outer(5)); end."#);
}

#[test] fn func_multiple_params() {
    parse_succeeds("program T; function Sum(a, b, c: Integer): Integer; begin Result := a + b + c; end; begin WriteLn(Sum(1, 2, 3)); end.");
}

#[test] fn func_exit_early() {
    parse_succeeds(r#"program T;
function Check(x: Integer): Integer;
begin if x > 10 then begin Result := 99; Exit; end; Result := x; end;
begin WriteLn(Check(5)); end."#);
}

// ═══════════════════════════════════════════════════════════
// BUILTINS (just parse, not execute)
// ═══════════════════════════════════════════════════════════

#[test] fn builtin_writeln()  { parse_succeeds("program T; begin WriteLn('hello'); end."); }
#[test] fn builtin_multi()    { parse_succeeds("program T; begin WriteLn('a', 'b', 'c'); end."); }
#[test] fn builtin_length()   { parse_succeeds("program T; begin WriteLn(Length('hello')); end."); }
#[test] fn builtin_concat()   { parse_succeeds("program T; begin WriteLn(Concat('a', 'b', 'c')); end."); }
#[test] fn builtin_inttostr() { parse_succeeds("program T; begin WriteLn(IntToStr(42)); end."); }
#[test] fn builtin_inc_dec()  { parse_succeeds("program T; var x: Integer; begin x := 5; Inc(x); Dec(x); end."); }
#[test] fn builtin_abs()      { parse_succeeds("program T; begin WriteLn(Abs(-5)); end."); }
#[test] fn builtin_sqrt()     { parse_succeeds("program T; begin WriteLn(Sqrt(16)); end."); }

// ═══════════════════════════════════════════════════════════
// STRINGS
// ═══════════════════════════════════════════════════════════

#[test] fn str_concat()    { parse_succeeds("program T; begin WriteLn('foo' + 'bar'); end."); }
#[test] fn str_escape()    { parse_succeeds("program T; begin WriteLn('it''s'); end."); }

// ═══════════════════════════════════════════════════════════
// ARRAYS
// ═══════════════════════════════════════════════════════════

#[test] fn array_literal()  { parse_succeeds("program T; var a: array of Integer; begin a := [10, 20, 30]; WriteLn(a[0]); end."); }
#[test] fn array_assign()   { parse_succeeds("program T; var a: array of Integer; begin a := [1, 2, 3]; a[1] := 99; end."); }

// ═══════════════════════════════════════════════════════════
// CLASSES
// ═══════════════════════════════════════════════════════════

#[test] fn class_basic() {
    parse_succeeds(r#"program T;
type TFoo = class
  public
    FVal: Integer;
    constructor Create(V: Integer);
    function GetVal: Integer;
  end;
constructor TFoo.Create(V: Integer); begin FVal := V; end;
function TFoo.GetVal: Integer; begin Result := FVal; end;
begin end."#);
}

#[test] fn class_inheritance() {
    parse_succeeds(r#"program T;
type TBase = class
  public
    FX: Integer;
    constructor Create(X: Integer);
  end;
type TChild = class(TBase)
  public
    constructor Create(X: Integer);
  end;
constructor TBase.Create(X: Integer); begin FX := X; end;
constructor TChild.Create(X: Integer); begin inherited Create(X); end;
begin end."#);
}

#[test] fn class_with_methods() {
    // First test: just two method impls (known to work)
    parse_succeeds(r#"program T;
type TCalc = class
  public
    FVal: Integer;
    constructor Create(V: Integer);
    function GetVal: Integer;
  end;
constructor TCalc.Create(V: Integer); begin FVal := V; end;
function TCalc.GetVal: Integer; begin Result := FVal; end;
begin end."#);
}

#[test] fn class_three_methods() {
    // Test just the method implementations without the class type
    parse_succeeds(r#"program T;
function Foo: Integer; begin Result := 1; end;
function Bar: Integer; begin Result := 2; end;
function Baz: Integer; begin Result := 3; end;
begin end."#);
}

#[test] fn class_three_method_impls() {
    parse_succeeds(r#"program T;
type TCalc = class
  public
    FVal: Integer;
    constructor Create(V: Integer);
    function GetVal: Integer;
    function Add(X: Integer): Integer;
  end;
constructor TCalc.Create(V: Integer); begin FVal := V; end;
function TCalc.GetVal: Integer; begin Result := FVal; end;
function TCalc.Add(X: Integer): Integer; begin FVal := FVal + X; Result := FVal; end;
begin end."#);
}

// ═══════════════════════════════════════════════════════════
// ENUMS
// ═══════════════════════════════════════════════════════════

#[test] fn enum_basic() {
    parse_succeeds(r#"program T; type TColor = (Red, Green, Blue); var c: TColor; begin c := Green; end."#);
}

// ═══════════════════════════════════════════════════════════
// INTERFACES
// ═══════════════════════════════════════════════════════════

#[test] fn interface_basic() {
    parse_succeeds(r#"program T;
type IGreeter = interface
  function Greet: String;
end;
begin end."#);
}

// ═══════════════════════════════════════════════════════════
// PROGRAMS (complex, multi-feature)
// ═══════════════════════════════════════════════════════════

#[test] fn prog_fizzbuzz() {
    parse_succeeds(r#"program T; var i: Integer;
begin for i := 1 to 15 do begin
  if (i mod 15) = 0 then WriteLn('FizzBuzz')
  else if (i mod 3) = 0 then WriteLn('Fizz')
  else if (i mod 5) = 0 then WriteLn('Buzz')
  else WriteLn(i);
end; end."#);
}

#[test] fn prog_gcd() {
    parse_succeeds(r#"program T;
function GCD(a, b: Integer): Integer;
begin if b = 0 then Result := a else Result := GCD(b, a mod b); end;
begin WriteLn(GCD(48, 18)); end."#);
}

#[test] fn prog_is_prime() {
    parse_succeeds(r#"program T;
function IsPrime(n: Integer): Boolean;
var i: Integer;
begin
  if n < 2 then begin Result := false; Exit; end;
  i := 2;
  while i * i <= n do begin
    if (n mod i) = 0 then begin Result := false; Exit; end;
    i := i + 1;
  end;
  Result := true;
end;
begin
  WriteLn(IsPrime(7));
end."#);
}

#[test] fn prog_power() {
    parse_succeeds(r#"program T;
function Pow(base, exp: Integer): Integer;
var i, r: Integer;
begin r := 1; for i := 1 to exp do r := r * base; Result := r; end;
begin WriteLn(Pow(2, 10)); end."#);
}

#[test] fn prog_collatz() {
    parse_succeeds(r#"program T;
function CollatzSteps(n: Integer): Integer;
var steps: Integer;
begin
  steps := 0;
  while n <> 1 do begin
    if (n mod 2) = 0 then n := n div 2
    else n := 3 * n + 1;
    steps := steps + 1;
  end;
  Result := steps;
end;
begin WriteLn(CollatzSteps(6)); end."#);
}

// ═══════════════════════════════════════════════════════════
// NEW FEATURES (for..in, +=, lambdas, is/as, enums)
// ═══════════════════════════════════════════════════════════

#[test] fn forin_array() {
    parse_succeeds("program T; var x: Integer; begin for x in [10, 20, 30] do WriteLn(x); end.");
}

#[test] fn compound_string() {
    parse_succeeds("program T; var s: String; begin s := 'hello'; s += ' world'; end.");
}

#[test] fn is_check() {
    parse_succeeds(r#"program T;
type TFoo = class public constructor Create; end;
constructor TFoo.Create; begin end;
var f: TFoo;
begin f := TFoo.Create; if f is TFoo then WriteLn('yes'); end."#);
}

// ═══════════════════════════════════════════════════════════
// EDGE CASES
// ═══════════════════════════════════════════════════════════

#[test] fn edge_empty()        { parse_succeeds("program T; begin end."); }
#[test] fn edge_nested_blocks() { parse_succeeds("program T; begin begin WriteLn('a'); end; begin WriteLn('b'); end; end."); }
#[test] fn edge_many_locals()  { parse_succeeds("program T; var a,b,c,d,e: Integer; begin a:=1; b:=2; c:=3; d:=4; e:=5; WriteLn(a+b+c+d+e); end."); }

// ═══════════════════════════════════════════════════════════
// FEATURES FROM OLD PARSER — gap closure
// ═══════════════════════════════════════════════════════════

#[test] fn anonymous_proc() {
    parse_succeeds(r#"program T;
var p: procedure;
begin
  p := procedure begin WriteLn('hello'); end;
end."#);
}

#[test] fn anonymous_func() {
    parse_succeeds(r#"program T;
var f: function(a, b: Integer): Integer;
begin
  f := function(a, b: Integer): Integer begin Result := a + b; end;
end."#);
}

#[test] fn procedural_type_var() {
    parse_succeeds("program T; var p: procedure; begin end.");
}

#[test] fn function_type_var() {
    parse_succeeds("program T; var f: function(x: Integer): Integer; begin end.");
}

#[test] fn pointer_type() {
    parse_succeeds("program T; type PInteger = ^Integer; begin end.");
}

#[test] fn virtual_method() {
    parse_succeeds(r#"program T;
type TBase = class
  public
    function Greet: String; virtual;
  end;
type TChild = class(TBase)
  public
    function Greet: String; override;
  end;
begin end."#);
}

#[test] fn abstract_method() {
    parse_succeeds(r#"program T;
type TAbstract = class
  public
    function Calculate: Integer; virtual; abstract;
  end;
begin end."#);
}

#[test] fn address_of() {
    parse_succeeds("program T; var x: Integer; var p: Pointer; begin x := 42; p := @x; end.");
}

#[test] fn class_three_method_impls_fixed() {
    parse_succeeds(r#"program T;
type TCalc = class
  public
    FVal: Integer;
    constructor Create(V: Integer);
    function GetVal: Integer;
    function Add(X: Integer): Integer;
  end;
constructor TCalc.Create(V: Integer); begin FVal := V; end;
function TCalc.GetVal: Integer; begin Result := FVal; end;
function TCalc.Add(X: Integer): Integer; begin FVal := FVal + X; Result := FVal; end;
begin end."#);
}

#[test] fn operator_overload_class() {
    parse_succeeds(r#"program T;
type TVector = class
  public
    FX, FY: Real;
    constructor Create(X, Y: Real);
    class operator Add(a, b: TVector): TVector;
  end;
constructor TVector.Create(X, Y: Real); begin FX := X; FY := Y; end;
begin end."#);
}

#[test] fn operator_overload_record() {
    parse_succeeds(r#"program T;
type TPoint = record
    X, Y: Integer;
    operator Add(a, b: TPoint): TPoint;
  end;
begin end."#);
}

#[test] fn operator_overload_stored_in_modifiers() {
    let m = parse_ok(r#"program T;
type TFoo = class
  public
    class operator Add(a, b: TFoo): TFoo;
  end;
begin end."#);
    if let StmtKind::ClassDecl { members, .. } = &m.body[0].kind {
        if let StmtKind::FunctionDecl { name, modifiers, .. } = &members[0].kind {
            assert_eq!(name, "Add");
            assert!(modifiers.is_static, "class operator should be static");
            assert!(modifiers.extra.contains(&"operator".to_string()), "should have operator in extra");
        } else { panic!("expected FunctionDecl"); }
    } else { panic!("expected ClassDecl"); }
}

#[test] fn method_with_directives() {
    let m = parse_ok(r#"program T;
type TFoo = class
  public
    function DoSomething: Integer; virtual;
  end;
begin end."#);
    // The virtual directive should be stored in modifiers
    if let StmtKind::ClassDecl { members, .. } = &m.body[0].kind {
        if let StmtKind::FunctionDecl { modifiers, .. } = &members[0].kind {
            assert!(modifiers.is_virtual, "expected virtual modifier");
        }
    }
}

// ═══════════════════════════════════════════════════════════
// AST STRUCTURE CHECKS (verify correct node types)
// ═══════════════════════════════════════════════════════════

#[test] fn ast_program_name() {
    let m = parse_ok("program MyApp; begin end.");
    assert_eq!(m.name, "MyApp");
}

#[test] fn ast_var_decl() {
    let m = parse_ok("program T; var x: Integer; begin end.");
    assert!(m.body.iter().any(|s| matches!(s.kind, StmtKind::VarDecl { .. })));
}

#[test] fn ast_func_decl() {
    let m = parse_ok("program T; function Add(a, b: Integer): Integer; begin Result := a + b; end; begin end.");
    assert!(m.body.iter().any(|s| matches!(s.kind, StmtKind::FunctionDecl { .. })));
}

#[test] fn ast_class_decl() {
    let m = parse_ok(r#"program T; type TFoo = class public FVal: Integer; end; begin end."#);
    assert!(m.body.iter().any(|s| matches!(s.kind, StmtKind::ClassDecl { .. })));
}

#[test] fn ast_enum_decl() {
    let m = parse_ok("program T; type TColor = (Red, Green, Blue); begin end.");
    assert!(m.body.iter().any(|s| matches!(s.kind, StmtKind::EnumDecl { .. })));
}

#[test] fn ast_if_stmt() {
    let m = parse_ok("program T; begin if true then WriteLn('y'); end.");
    assert!(m.body.iter().any(|s| matches!(s.kind, StmtKind::If { .. })));
}

#[test] fn ast_for_loop() {
    let m = parse_ok("program T; var i: Integer; begin for i := 1 to 5 do WriteLn(i); end.");
    assert!(m.body.iter().any(|s| matches!(s.kind, StmtKind::For { .. })));
}

#[test] fn ast_while_loop() {
    let m = parse_ok("program T; begin while true do break; end.");
    assert!(m.body.iter().any(|s| matches!(s.kind, StmtKind::While { .. })));
}

#[test] fn ast_assign() {
    let m = parse_ok("program T; var x: Integer; begin x := 42; end.");
    assert!(m.body.iter().any(|s| matches!(s.kind, StmtKind::Assign { .. })));
}

#[test] fn ast_compound_assign() {
    let m = parse_ok("program T; var x: Integer; begin x := 0; x += 5; end.");
    assert!(m.body.iter().any(|s| matches!(s.kind, StmtKind::CompoundAssign { .. })));
}

#[test] fn ast_expression_precedence() {
    let m = parse_ok("program T; begin WriteLn(2 + 3 * 4); end.");
    // Should parse as 2 + (3 * 4), not (2 + 3) * 4
    if let StmtKind::Expr(ref expr) = m.body[0].kind {
        if let ExprKind::Call { ref args, .. } = expr.kind {
            if let ExprKind::Binary { ref op, ref right, .. } = args[0].kind {
                assert_eq!(*op, BinOp::Add);
                assert!(matches!(right.kind, ExprKind::Binary { op: BinOp::Mul, .. }));
            } else { panic!("expected Binary"); }
        }
    }
}

#[test] fn ast_multi_const_body_count() {
    let m = parse_ok("program T; const A = 10; B = 20; begin WriteLn(A + B); end.");
    // Should have: Block([VarDecl A, VarDecl B]) + Expr(Call WriteLn)
    // or: VarDecl A, VarDecl B, Expr(Call WriteLn)
    let var_decl_count = m.body.iter().filter(|s| matches!(s.kind, StmtKind::VarDecl { .. })).count();
    let block_count = m.body.iter().filter(|s| matches!(s.kind, StmtKind::Block(_))).count();
    eprintln!("body len={}, var_decls={}, blocks={}", m.body.len(), var_decl_count, block_count);
    for (i, s) in m.body.iter().enumerate() {
        eprintln!("  [{}] {:?}", i, std::mem::discriminant(&s.kind));
    }
    // The const section must produce both A and B somehow
    assert!(var_decl_count >= 2 || block_count >= 1, "expected both A and B in AST");
}

#[test] fn vb_hello_world_parse() {
    let g = vb_grammar();
    let tokens = vybe_parser_generic::lexer::tokenize(
        "Module Program\n    Sub Main()\n        Console.WriteLine(\"Hello\")\n    End Sub\nEnd Module\n",
        &g.lexer, &g.language.statement_terminator, false, false,
    );
    eprintln!("Tokens:");
    for (i, t) in tokens.iter().enumerate() {
        eprintln!("  [{:3}] {:?} '{}'", i, t.kind, t.text);
    }
    let result = vybe_parser_generic::parser::parse(&tokens, &g);
    match result {
        Ok(m) => eprintln!("Parsed OK: {} body items", m.body.len()),
        Err(e) => panic!("Parse failed: {}", e),
    }
}

fn vb_grammar() -> vybe_parser_generic::grammar::GrammarDef {
    use vybe_parser_generic::grammar::*;
    GrammarDef {
        language: LanguageSpec { name: "vb".into(), case_sensitive: false, statement_terminator: Terminator::Newline, indentation_based: false, expression_language: false },
        lexer: LexerSpec {
            comment_line: vec!["'".into()],
            comment_block: Vec::new(),
            string_delimiters: vec!["\"".into()],
            string_escape: Some("\"\"".into()),
            triple_string: Vec::new(), string_prefixes: Vec::new(), interpolation: None, template_string: None,
            char_prefix: None, hex_prefix: None,
            keywords: vec!["module".into(),"end".into(),"sub".into(),"function".into(),"dim".into(),"as".into(),"if".into(),"then".into(),"else".into(),"elseif".into(),"for".into(),"to".into(),"next".into(),"each".into(),"in".into(),"while".into(),"do".into(),"loop".into(),"select".into(),"case".into(),"class".into(),"public".into(),"private".into(),"return".into(),"true".into(),"false".into(),"nothing".into(),"and".into(),"andalso".into(),"or".into(),"orelse".into(),"not".into(),"me".into(),"new".into(),"console".into(),"writeline".into(),"integer".into(),"string".into(),"boolean".into(),"double".into(),"object".into()],
            operators: vec!["<>".into(),"<=".into(),">=".into(),"+=".into(),"-=".into(),"*=".into(),"/=".into(),"&=".into(),"+".into(),"-".into(),"*".into(),"/".into(),"\\".into(),"^".into(),"&".into(),"=".into(),"<".into(),">".into(),"(".into(),")".into(),"[".into(),"]".into(),".".into(),",".into(),":".into()],
        },
        operators: OperatorTable {
            prefix: vec!["not".into(), "-".into()],
            postfix: Vec::new(),
            infix: vec![
                InfixLevel { precedence: 1, ops: vec!["or".into(),"orelse".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 2, ops: vec!["and".into(),"andalso".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 3, ops: vec!["=".into(),"<>".into(),"<".into(),">".into(),"<=".into(),">=".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 4, ops: vec!["&".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 5, ops: vec!["+".into(),"-".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 6, ops: vec!["*".into(),"/".into(),"\\".into()], assoc: Assoc::Left },
            ],
        },
        blocks: BlockSpec { open: "SUB_BLOCK".into(), close: "end".into(), prefix: None, close_with_kind: true },
        types: TypeSpec { position: TypePosition::After, separator: Some("as".into()), return_separator: Some("as".into()) },
        statements: Vec::new(), declarations: Vec::new(),
        expressions: ExpressionSpec { member_access: Some(".".into()), optional_chain: None, index_open: Some("(".into()), index_close: Some(")".into()), call_open: Some("(".into()), call_close: Some(")".into()), deref: None, primary_forms: Vec::new() },
        params: ParamSpec { open: "(".into(), close: ")".into(), separator: ",".into(), name_type_sep: Some("as".into()), type_position: TypePosition::After, default_value: Some("=".into()), rest_prefix: None, kwargs_prefix: None, multi_name: false, multi_name_sep: None, pass_by: std::collections::HashMap::new() },
        assignment: AssignmentSpec { operator: Some("=".into()), compound: [("+=".into(),"Add".into()),("-=".into(),"Sub".into()),("*=".into(),"Mul".into()),("/=".into(),"Div".into()),("&=".into(),"Concat".into())].into_iter().collect(), walrus: None },
        program: ProgramSpec { header: None, uses: None, body: None },
    }
}
