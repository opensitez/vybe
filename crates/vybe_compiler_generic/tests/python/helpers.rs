use std::rc::Rc;
use std::cell::RefCell;
use vybe_bytecode::{VM, Value, HostContext};
use vybe_parser_generic::grammar::*;
use vybe_parser_generic::profile::LanguageProfile;
use vybe_parser_generic::lexer::tokenize;
use vybe_parser_generic::parser::parse;

pub fn run_prints(src: &str) -> Vec<String> { run(src) }

pub fn compile(src: &str) -> Vec<vybe_bytecode::Chunk> {
    let g = grammar();
    let tokens = tokenize(src, &g.lexer, &g.language.statement_terminator, true, true);
    let module = parse(&tokens, &g).expect("parse failed");
    vybe_compiler_generic::Compiler::with_profile(LanguageProfile::python())
        .compile(&module).expect("compile failed")
}

pub fn run(src: &str) -> Vec<String> {
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
    let tokens = tokenize(src, &g.lexer, &g.language.statement_terminator, true, true);
    let module = parse(&tokens, &g).expect("parse failed");
    let chunks = vybe_compiler_generic::Compiler::with_profile(LanguageProfile::python())
        .compile(&module).expect("compile failed");
    vm.run(chunks).expect("run failed");
    output.borrow().clone()
}

/// Compile only — no VM execution. Just verify parse + compile doesn't crash.
pub fn compile_ok(src: &str) {
    let g = grammar();
    let tokens = tokenize(src, &g.lexer, &g.language.statement_terminator, true, true);
    let module = parse(&tokens, &g).expect("parse failed");
    let _chunks = vybe_compiler_generic::Compiler::with_profile(LanguageProfile::python())
        .compile(&module).expect("compile failed");
}

fn grammar() -> GrammarDef {
    GrammarDef {
        language: LanguageSpec { name: "python".into(), case_sensitive: true, statement_terminator: Terminator::Newline, indentation_based: true, expression_language: false },
        lexer: LexerSpec {
            comment_line: vec!["#".into()],
            comment_block: Vec::new(),
            string_delimiters: vec!["'".into(), "\"".into()],
            string_escape: Some("\\".into()),
            triple_string: vec!["'''".into(), "\"\"\"".into()],
            string_prefixes: vec!["f".into(), "r".into(), "b".into(), "F".into(), "R".into(), "B".into()],
            interpolation: Some(("{".into(), "}".into())),
            template_string: None,
            char_prefix: None, hex_prefix: None,
            keywords: vec![
                "def".into(), "class".into(), "if".into(), "elif".into(), "else".into(),
                "for".into(), "while".into(), "break".into(), "continue".into(), "return".into(),
                "pass".into(), "yield".into(),
                "try".into(), "except".into(), "finally".into(), "raise".into(),
                "with".into(), "as".into(),
                "import".into(), "from".into(),
                "and".into(), "or".into(), "not".into(), "in".into(), "is".into(),
                "lambda".into(), "global".into(), "nonlocal".into(), "assert".into(), "del".into(),
                "True".into(), "False".into(), "None".into(),
                "async".into(), "await".into(),
            ],
            operators: vec![
                "**=".into(), "//=".into(), ">>=".into(), "<<=".into(),
                "**".into(), "//".into(), ">>".into(), "<<".into(),
                "+=".into(), "-=".into(), "*=".into(), "/=".into(), "%=".into(),
                "&=".into(), "|=".into(), "^=".into(),
                "==".into(), "!=".into(), "<=".into(), ">=".into(),
                "->".into(), ":=".into(),
                "+".into(), "-".into(), "*".into(), "/".into(), "%".into(),
                "&".into(), "|".into(), "^".into(), "~".into(),
                "<".into(), ">".into(), "=".into(),
                "(".into(), ")".into(), "[".into(), "]".into(), "{".into(), "}".into(),
                ".".into(), ",".into(), ";".into(), ":".into(), "@".into(),
            ],
        },
        operators: OperatorTable {
            prefix: vec!["not".into(), "-".into(), "+".into(), "~".into()],
            postfix: Vec::new(),
            infix: vec![
                InfixLevel { precedence: 1, ops: vec!["or".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 2, ops: vec!["and".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 3, ops: vec!["==".into(),"!=".into(),"<".into(),">".into(),"<=".into(),">=".into(),"in".into(),"is".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 4, ops: vec!["|".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 5, ops: vec!["^".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 6, ops: vec!["&".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 7, ops: vec!["<<".into(), ">>".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 8, ops: vec!["+".into(), "-".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 9, ops: vec!["*".into(), "/".into(), "//".into(), "%".into()], assoc: Assoc::Left },
                InfixLevel { precedence: 10, ops: vec!["**".into()], assoc: Assoc::Right },
            ],
        },
        blocks: BlockSpec { open: "INDENT".into(), close: "DEDENT".into(), prefix: Some(":".into()), close_with_kind: false },
        types: TypeSpec { position: TypePosition::After, separator: Some(":".into()), return_separator: Some("->".into()) },
        statements: Vec::new(), declarations: Vec::new(),
        expressions: ExpressionSpec {
            member_access: Some(".".into()), optional_chain: None,
            index_open: Some("[".into()), index_close: Some("]".into()),
            call_open: Some("(".into()), call_close: Some(")".into()),
            deref: None, primary_forms: Vec::new(),
        },
        params: ParamSpec {
            open: "(".into(), close: ")".into(), separator: ",".into(),
            name_type_sep: Some(":".into()), type_position: TypePosition::After,
            default_value: Some("=".into()),
            rest_prefix: Some("*".into()), kwargs_prefix: Some("**".into()),
            multi_name: false, multi_name_sep: None,
            pass_by: std::collections::HashMap::new(),
        },
        assignment: AssignmentSpec {
            operator: Some("=".into()),
            compound: [
                ("+=".into(),"Add".into()), ("-=".into(),"Sub".into()),
                ("*=".into(),"Mul".into()), ("/=".into(),"Div".into()),
                ("//=".into(),"IDiv".into()), ("%=".into(),"Mod".into()),
                ("**=".into(),"Pow".into()),
            ].into_iter().collect(),
            walrus: Some(":=".into()),
        },
        program: ProgramSpec { header: None, uses: None, body: None },
    }
}
