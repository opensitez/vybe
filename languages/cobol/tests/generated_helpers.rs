use vybe_ast::{ExprKind, Module, StmtKind};

const SOURCE: &str = "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC X(80).\n01 B PIC X(80).\nPROCEDURE DIVISION.\nACCEPT A FROM COMMAND-LINE.\nACCEPT B FROM COMMAND-LINE.\nSTOP RUN.\n";

fn calls(module: &Module, target: &str) -> usize {
    let mut count = 0;
    for statement in &module.body {
        statement.walk_exprs(&mut |expression| {
            if let ExprKind::Call { callee, .. } = &expression.kind {
                if matches!(&callee.kind, ExprKind::Ident(name) if name == target) {
                    count += 1;
                }
            }
        });
    }
    count
}

#[test]
fn command_line_sites_use_one_provider_and_keep_one_host_read_in_its_body() {
    let module = vybe_language_cobol::parse(SOURCE).unwrap();
    let providers: Vec<_> = module
        .body
        .iter()
        .filter_map(|statement| match &statement.kind {
            StmtKind::FunctionDecl { name, .. }
                if name.starts_with("<vybe:cobol.command-line:") =>
            {
                Some(name)
            }
            _ => None,
        })
        .collect();
    assert_eq!(providers.len(), 1);
    assert_eq!(calls(&module, providers[0]), 2);
    assert_eq!(calls(&module, "__cobol_args"), 1);
    // Publication belongs to this parse, never a previous module's requests.
    let no_accept = vybe_language_cobol::parse(
        &SOURCE
            .replace("ACCEPT A FROM COMMAND-LINE.", "CONTINUE.")
            .replace("ACCEPT B FROM COMMAND-LINE.", "CONTINUE."),
    )
    .unwrap();
    assert_eq!(calls(&no_accept, "__cobol_args"), 0);
    assert!(!no_accept.body.iter().any(|statement| matches!(&statement.kind, StmtKind::FunctionDecl { name, .. } if name.starts_with("<vybe:cobol.command-line:"))));
}

#[test]
fn possible_shadowing_keeps_both_original_inline_host_reads() {
    // Conservative even in a literal: the source can name a dependency through
    // dynamic operations. No lexical dependency is silently hoisted.
    let source = SOURCE.replace("STOP RUN.", "DISPLAY '__COBOL_ARGS'.\nSTOP RUN.");
    let module = vybe_language_cobol::parse(&source).unwrap();
    assert_eq!(calls(&module, "__cobol_args"), 2);
    assert!(!module.body.iter().any(|statement| matches!(&statement.kind, StmtKind::FunctionDecl { name, .. } if name.starts_with("<vybe:cobol.command-line:"))));
}

#[test]
fn environment_literals_share_only_the_literal_provider() {
    let source = SOURCE.replace("COMMAND-LINE", "ENVIRONMENT 'HOME'");
    let module = vybe_language_cobol::parse(&source).unwrap();
    assert_eq!(calls(&module, "__cobol_env_pairs"), 1);
    assert_eq!(calls(&module, "<vybe:cobol.environment-literal:0>"), 2);
    let dynamic =
        vybe_language_cobol::parse(&SOURCE.replace("COMMAND-LINE", "ENVIRONMENT NAME A")).unwrap();
    assert_eq!(calls(&dynamic, "__cobol_env_pairs"), 2);
    assert_eq!(calls(&dynamic, "<vybe:cobol.environment-literal:0>"), 0);
}
