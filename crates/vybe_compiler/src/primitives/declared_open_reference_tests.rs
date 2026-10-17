use super::*;

#[cfg(test)]
mod declared_open_reference_tests {
    use super::*;
    fn make_compiler(fold: bool) -> Compiler {
        let profile=vybe_runtime::profile::parse_profile(&format!("[info]\nname = 'reference-probe'\n[compiler]\nreference_creates_local = true\ncase_sensitive = {}\n",!fold)).unwrap();
        let mut compiler = Compiler::with_profile(profile);
        compiler.scopes.push(Scope::new_function(
            fold.then_some(vybe_ast::CaseAlphabet::Ascii),
        ));
        compiler.current_func_name = Some("probe".into());
        compiler
    }
    fn local_gets(chunk: &Chunk, slot: u16) -> usize {
        let mut ip = 0;
        let mut count = 0;
        while ip < chunk.code.len() {
            let op = Op::from_code_at(&chunk.code, ip);
            if op == Op::LOCAL_GET
                && u16::from_be_bytes(chunk.code[ip + 4..ip + 6].try_into().unwrap()) == slot
            {
                count += 1;
            }
            ip += 4 + op.operand_format().size_in(&chunk.code, ip + 4);
        }
        count
    }

    #[test]
    fn declared_open_reference_ignores_null_local_alias_for_raw_and_canonical_names() {
        for (raw, fold) in [(true, false), (true, true), (false, true)] {
            let mut compiler = make_compiler(fold);
            let name = "OPEN";
            let slot = compiler.define_source_local(name);
            let opened = if raw {
                name.to_string()
            } else {
                compiler.canon(name)
            };
            compiler.scope_mut().close(&[opened]);
            let key = compiler.variable_global_binding_key(name);
            compiler.compile_ident_reference(name);
            assert_eq!(
                local_gets(&compiler.chunks[0], slot),
                0,
                "must not box/read the local alias"
            );
            assert!(
                compiler.chunks[0]
                    .constants
                    .iter()
                    .any(|i| matches!(i, vybe_runtime::Value::String(s) if s.as_ref() == key))
            );
        }
    }

    #[test]
    fn local_and_undeclared_reference_storage_remains_local() {
        let mut compiler = make_compiler(false);
        compiler.scope_mut().close(&[]);
        let slot = compiler.define_source_local("local");
        compiler.compile_ident_reference("local");
        assert!(local_gets(&compiler.chunks[0], slot) > 0);
        assert!(
            !compiler.chunks[0]
                .constants
                .iter()
                .any(|i| matches!(i, vybe_runtime::Value::String(s) if s.as_ref() == "local"))
        );
        static VARIABLES: vybe_runtime::registry::VariableNamespace =
            vybe_runtime::registry::VariableNamespace {
                body: |name| name.strip_prefix('$').unwrap_or(name),
                spell: |name| format!("${name}"),
                global_key: |_, _| None,
            };
        compiler.variable_namespace = Some(&VARIABLES);
        compiler.compile_ident_reference("$fresh");
        assert!(compiler.scope().resolve("$fresh").is_some());
    }

    #[test]
    fn declared_open_undeclared_reference_does_not_create_source_local() {
        let mut compiler = make_compiler(false);
        compiler.scope_mut().close(&["outer".into()]);
        compiler.compile_ident_reference("outer");
        assert!(compiler.scope().resolve("outer").is_none());
        assert!(
            compiler.chunks[0]
                .constants
                .iter()
                .any(|i| matches!(i, vybe_runtime::Value::String(s) if s.as_ref() == "outer"))
        );
    }

    #[test]
    fn static_and_shared_environment_reference_precedence_is_unchanged() {
        let mut compiler = make_compiler(false);
        let slot = compiler.define_source_local("outer");
        compiler.scope_mut().close(&["outer".into()]);
        compiler.static_local_bindings.push(
            [(
                compiler.canon("outer"),
                StaticLocalBinding {
                    global_name: "private-static-state".into(),
                    init_flag_name: "private-static-init".into(),
                    type_hint: None,
                    declaration_scope_depth: compiler.scopes.len(),
                },
            )]
            .into_iter()
            .collect(),
        );
        compiler.compile_ident_reference("outer");
        assert_eq!(local_gets(&compiler.chunks[0], slot), 0);
        assert!(compiler.chunks[0].constants.iter().any(
            |i| matches!(i, vybe_runtime::Value::String(s) if s.as_ref() == "private-static-state")
        ));
        assert!(
            !compiler.chunks[0]
                .constants
                .iter()
                .any(|i| matches!(i, vybe_runtime::Value::String(s) if s.as_ref() == "outer"))
        );
        let mut compiler = make_compiler(false);
        let env = compiler.define_source_local("env");
        let slot = compiler.define_source_local("outer");
        compiler.scope_mut().close(&["outer".into()]);
        compiler.shared_env_slot = Some(env);
        compiler.shared_env_names = vec!["outer".into()];
        compiler.compile_ident_reference("outer");
        assert!(local_gets(&compiler.chunks[0], slot) > 0);
        assert!(local_gets(&compiler.chunks[0], env) > 0);
        assert!(
            !compiler.chunks[0]
                .constants
                .iter()
                .any(|i| matches!(i, vybe_runtime::Value::String(s) if s.as_ref() == "outer"))
        );
    }
}
