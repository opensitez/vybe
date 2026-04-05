//! vybeg — Run any language through the generic parser + compiler pipeline.
//!
//! Usage: vybeg <file.pas|file.py|file.js|...>
//!
//! Compare with vybec (language-specific compilers) to validate parity.

use std::path::Path;
use std::rc::Rc;
use std::cell::RefCell;
use std::collections::HashMap;
use vybe_bytecode::{VM, Value, HostContext};
use vybe_parser_generic::grammar::*;
use vybe_parser_generic::lexer::tokenize;
use vybe_parser_generic::parser::parse;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: vybeg <file>");
        eprintln!("Supported: .pas .py .js .rb .php .dart .cs .vb .cob");
        std::process::exit(1);
    }

    let path = Path::new(&args[1]);
    let source = match std::fs::read_to_string(path) {
        Ok(s) => s,
        Err(e) => { eprintln!("Cannot read {}: {}", path.display(), e); std::process::exit(1); }
    };

    let ext = path.extension().and_then(|e| e.to_str()).unwrap_or("");
    let grammar = match build_grammar(ext) {
        Some(g) => g,
        None => { eprintln!("Unsupported extension: .{}", ext); std::process::exit(1); }
    };

    let is_indent = grammar.language.indentation_based;
    let case_sensitive = grammar.language.case_sensitive;

    // Parse
    let tokens = tokenize(&source, &grammar.lexer, &grammar.language.statement_terminator, is_indent, case_sensitive);
    let module = match parse(&tokens, &grammar) {
        Ok(m) => m,
        Err(e) => { eprintln!("Parse error: {}", e); std::process::exit(1); }
    };

    // Compile with language profile
    use vybe_parser_generic::profile::LanguageProfile;
    let profile = match ext.to_lowercase().as_str() {
        "pas" | "pp" | "dpr" | "lpr" => LanguageProfile::pascal(),
        "py" | "pyw" => LanguageProfile::python(),
        "vb" => LanguageProfile::vb(),
        "js" | "mjs" => LanguageProfile::javascript(),
        _ => LanguageProfile::vb(), // fallback
    };
    let chunks = match vybe_compiler_generic::Compiler::with_profile(profile).compile(&module) {
        Ok(c) => c,
        Err(e) => { eprintln!("Compile error: {}", e); std::process::exit(1); }
    };

    // Run
    let mut vm = VM::new();
    let queue = Rc::new(RefCell::new(vybe_host::SideEffectQueue::new()));
    let gui = vybe_host::register_all_with_gui(&mut vm, queue.clone());

    match vm.run(chunks) {
        Ok(_) => {}
        Err(e) => { eprintln!("Runtime error: {}", e); std::process::exit(1); }
    }

    // If the program created forms, launch the GUI
    if gui.borrow().should_run {
        vybe_cli::runner::launch_vm_form(vm, queue, gui, None);
    }
}

fn build_grammar(ext: &str) -> Option<GrammarDef> {
    match ext.to_lowercase().as_str() {
        "pas" | "pp" | "dpr" | "lpr" => Some(pascal_grammar()),
        "py" | "pyw" => Some(python_grammar()),
        "vb" => Some(vb_grammar()),
        _ => None,
    }
}

fn pascal_grammar() -> GrammarDef {
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

fn python_grammar() -> GrammarDef {
    GrammarDef {
        language: LanguageSpec { name: "python".into(), case_sensitive: true, statement_terminator: Terminator::Newline, indentation_based: true, expression_language: false },
        lexer: LexerSpec {
            comment_line: vec!["#".into()],
            comment_block: Vec::new(),
            string_delimiters: vec!["'".into(), "\"".into()],
            string_escape: Some("\\".into()),
            triple_string: vec!["'''".into(), "\"\"\"".into()],
            string_prefixes: vec!["f".into(), "r".into(), "b".into(), "fr".into(), "rb".into(), "F".into(), "R".into(), "B".into()],
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
                "match".into(), "case".into(),
            ],
            operators: vec![
                "**=".into(), "//=".into(), ">>=".into(), "<<=".into(),
                "**".into(), "//".into(), ">>".into(), "<<".into(),
                "+=".into(), "-=".into(), "*=".into(), "/=".into(), "%=".into(),
                "&=".into(), "|=".into(), "^=".into(), "@=".into(),
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
            pass_by: HashMap::new(),
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

fn vb_grammar() -> GrammarDef {
    GrammarDef {
        language: LanguageSpec {
            name: "vb".into(), case_sensitive: false,
            statement_terminator: Terminator::Newline, indentation_based: false, expression_language: false,
        },
        lexer: LexerSpec {
            comment_line: vec!["'".into()],
            comment_block: Vec::new(),
            string_delimiters: vec!["\"".into()],
            string_escape: Some("\"\"".into()),
            triple_string: Vec::new(), string_prefixes: Vec::new(),
            interpolation: None, template_string: None,
            char_prefix: None, hex_prefix: None,
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
        statements: Vec::new(), declarations: Vec::new(),
        expressions: ExpressionSpec {
            member_access: Some(".".into()), optional_chain: None,
            index_open: Some("(".into()), index_close: Some(")".into()),
            call_open: Some("(".into()), call_close: Some(")".into()),
            deref: None, primary_forms: Vec::new(),
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
