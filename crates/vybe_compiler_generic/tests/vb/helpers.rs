use std::rc::Rc;
use std::cell::RefCell;
use vybe_bytecode::{VM, Value, HostContext};
use vybe_parser_generic::grammar::*;
use vybe_parser_generic::profile::LanguageProfile;
use vybe_parser_generic::lexer::tokenize;
use vybe_parser_generic::parser::parse;

pub fn run_vb(src: &str) -> Vec<String> {
    let g = grammar();
    let mut vm = VM::new();
    let output: Rc<RefCell<Vec<String>>> = Rc::new(RefCell::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.borrow_mut().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    let tokens = tokenize(src, &g.lexer, &g.language.statement_terminator, false, false);
    let module = parse(&tokens, &g).expect("parse failed");
    let chunks = vybe_compiler_generic::Compiler::with_profile(LanguageProfile::vb())
        .compile(&module).expect("compile failed");
    vm.run(chunks).expect("run failed");
    output.borrow().clone()
}

fn grammar() -> GrammarDef {
    GrammarDef {
        language: LanguageSpec {
            name: "vb".into(),
            case_sensitive: false,
            statement_terminator: Terminator::Newline,
            indentation_based: false,
            expression_language: false,
        },
        lexer: LexerSpec {
            comment_line: vec!["'".into()],
            comment_block: Vec::new(),
            string_delimiters: vec!["\"".into()],
            string_escape: Some("\"\"".into()),
            triple_string: Vec::new(),
            string_prefixes: Vec::new(),
            interpolation: None,
            template_string: None,
            char_prefix: None,
            hex_prefix: None,
            keywords: vec![
                "module".into(), "end".into(), "sub".into(), "function".into(),
                "dim".into(), "as".into(), "new".into(), "nothing".into(),
                "if".into(), "then".into(), "else".into(), "elseif".into(), "end if".into(),
                "for".into(), "to".into(), "step".into(), "next".into(), "each".into(), "in".into(),
                "while".into(), "do".into(), "loop".into(), "until".into(), "wend".into(),
                "select".into(), "case".into(),
                "class".into(), "inherits".into(), "implements".into(),
                "public".into(), "private".into(), "protected".into(), "friend".into(), "shared".into(),
                "property".into(), "get".into(), "set".into(),
                "try".into(), "catch".into(), "finally".into(), "throw".into(),
                "return".into(), "exit".into(),
                "imports".into(), "namespace".into(),
                "interface".into(), "enum".into(), "structure".into(),
                "overridable".into(), "overrides".into(), "mustoverride".into(), "notoverridable".into(),
                "and".into(), "andalso".into(), "or".into(), "orelse".into(), "not".into(), "xor".into(), "mod".into(),
                "is".into(), "isnot".into(), "like".into(), "typeof".into(),
                "true".into(), "false".into(),
                "integer".into(), "string".into(), "double".into(), "boolean".into(), "object".into(),
                "date".into(), "byte".into(), "long".into(), "short".into(), "single".into(), "char".into(), "decimal".into(),
                "me".into(), "mybase".into(), "myclass".into(),
                "withevents".into(), "handles".into(), "addhandler".into(), "removehandler".into(), "raiseevent".into(),
                "event".into(), "delegate".into(),
                "byval".into(), "byref".into(), "optional".into(), "paramarray".into(),
                "readonly".into(), "writeonly".into(), "const".into(),
                "static".into(), "shadows".into(), "overloads".into(),
                "cbool".into(), "cbyte".into(), "cchar".into(), "cdate".into(), "cdbl".into(),
                "cdec".into(), "cint".into(), "clng".into(), "cobj".into(), "cshort".into(),
                "csng".into(), "cstr".into(), "ctype".into(), "directcast".into(), "trycast".into(),
                "of".into(), "with".into(),
                "console".into(), "writeline".into(), "readline".into(),
                "msgbox".into(), "messagebox".into(),
            ],
            operators: vec![
                "<>".into(), "<=".into(), ">=".into(),
                "+=".into(), "-=".into(), "*=".into(), "/=".into(), "\\=".into(), "&=".into(),
                "+".into(), "-".into(), "*".into(), "/".into(), "\\".into(), "^".into(), "&".into(),
                "=".into(), "<".into(), ">".into(),
                "(".into(), ")".into(), "[".into(), "]".into(), "{".into(), "}".into(),
                ".".into(), ",".into(), ":".into(),
            ],
        },
        operators: OperatorTable {
            prefix: vec!["not".into(), "-".into(), "+".into()],
            postfix: Vec::new(),
            infix: vec![
                InfixLevel { precedence: 1, ops: vec!["orelse".into(), "or".into(), "xor".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 2, ops: vec!["andalso".into(), "and".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 3, ops: vec!["=".into(), "<>".into(), "<".into(), ">".into(), "<=".into(), ">=".into(), "is".into(), "isnot".into(), "like".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 4, ops: vec!["&".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 5, ops: vec!["+".into(), "-".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 6, ops: vec!["*".into(), "/".into(), "\\".into(), "mod".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 7, ops: vec!["^".into()], assoc: Assoc::Right },
            ],
        },
        blocks: BlockSpec { open: "SUB_BLOCK".into(), close: "end".into(), prefix: None, close_with_kind: true },
        types: TypeSpec { position: TypePosition::After, separator: Some("as".into()), return_separator: Some("as".into()) },
        statements: Vec::new(),
        declarations: Vec::new(),
        expressions: ExpressionSpec {
            member_access: Some(".".into()),
            optional_chain: None,
            index_open: Some("(".into()),  // VB uses () for both calls and indexing
            index_close: Some(")".into()),
            call_open: Some("(".into()),
            call_close: Some(")".into()),
            deref: None,
            primary_forms: Vec::new(),
        },
        params: ParamSpec {
            open: "(".into(), close: ")".into(), separator: ",".into(),
            name_type_sep: Some("as".into()), type_position: TypePosition::After,
            default_value: Some("=".into()),
            rest_prefix: None, kwargs_prefix: None,
            multi_name: false, multi_name_sep: None,
            pass_by: [("byval".into(), "value".into()), ("byref".into(), "ref".into())].into_iter().collect(),
        },
        assignment: AssignmentSpec {
            operator: Some("=".into()),
            compound: [
                ("+=".into(), "Add".into()), ("-=".into(), "Sub".into()),
                ("*=".into(), "Mul".into()), ("/=".into(), "Div".into()),
                ("&=".into(), "Concat".into()),
            ].into_iter().collect(),
            walrus: None,
        },
        program: ProgramSpec { header: None, uses: None, body: None },
    }
}
