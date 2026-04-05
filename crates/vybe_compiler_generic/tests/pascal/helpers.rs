use std::rc::Rc;
use std::cell::RefCell;
use vybe_bytecode::{VM, Value, HostContext};
use vybe_parser_generic::grammar::*;
use vybe_parser_generic::lexer::tokenize;
use vybe_parser_generic::parser::parse;

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
    let tokens = tokenize(src, &g.lexer, &g.language.statement_terminator, false, false);
    let module = parse(&tokens, &g).expect("parse failed");
    let chunks = vybe_compiler_generic::Compiler::new().compile(&module).expect("compile failed");
    vm.run(chunks).expect("run failed");
    output.borrow().clone()
}

fn grammar() -> GrammarDef {
    GrammarDef {
        language: LanguageSpec { name: "pascal".into(), case_sensitive: false, statement_terminator: Terminator::Char(';'), indentation_based: false, expression_language: false },
        lexer: LexerSpec {
            comment_line: vec!["//".into()],
            comment_block: vec![("{".into(), "}".into()), ("(*".into(), "*)".into())],
            string_delimiters: vec!["'".into()],
            string_escape: Some("''".into()),
            triple_string: Vec::new(), string_prefixes: Vec::new(), interpolation: None, template_string: None,
            char_prefix: Some("#".into()), hex_prefix: Some("$".into()),
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
                "new".into(),"dispose".into(),"writeln".into(),"write".into(),
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
        statements: Vec::new(), declarations: Vec::new(),
        expressions: ExpressionSpec {
            member_access: Some(".".into()), optional_chain: None,
            index_open: Some("[".into()), index_close: Some("]".into()),
            call_open: Some("(".into()), call_close: Some(")".into()),
            deref: Some("^".into()), primary_forms: Vec::new(),
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
