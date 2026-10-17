//! Explicit generated-provider ownership. Never deduplicate guest lambdas.
use super::*;
use std::{cell::Cell, rc::Rc};

struct Provider {
    name: String,
    requested: Cell<bool>,
}

/// Shared only by scopes belonging to the same parsed module. Independently
/// created OO method contexts retain the original inline producer.
#[derive(Clone, Default)]
pub(super) struct Requests {
    command_line: Option<Rc<Provider>>,
    environment: Option<Rc<Provider>>,
}

fn contains_folded(source: &str, text: &str) -> bool {
    source
        .as_bytes()
        .windows(text.len())
        .any(|part| part.eq_ignore_ascii_case(text.as_bytes()))
}

impl Requests {
    pub fn for_source(source: &str) -> Self {
        fn provider(
            source: &str,
            keyword: &str,
            dependency: &str,
            kind: &str,
        ) -> Option<Rc<Provider>> {
            if !contains_folded(source, keyword) || contains_folded(source, dependency) {
                return None;
            }
            let name = (0usize..)
                .map(|suffix| format!("<vybe:cobol.{kind}:{suffix}>"))
                .find(|name| !contains_folded(source, name))
                .unwrap();
            Some(Rc::new(Provider {
                name,
                requested: Cell::new(false),
            }))
        }
        Self {
            command_line: provider(source, "COMMAND-LINE", "__cobol_args", "command-line"),
            environment: provider(
                source,
                "ENVIRONMENT",
                "__cobol_env_pairs",
                "environment-literal",
            ),
        }
    }

    pub fn command_line(&self) -> Expression {
        let Some(provider) = &self.command_line else {
            return cobol_join_args_inline();
        };
        provider.requested.set(true);
        // The old IIFE has no arguments or captures. Keep its host read inside
        // the shared body, in its original order; no input effects move here.
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(&provider.name)),
            args: Vec::new(),
            optional: false,
        })
    }

    pub fn environment(&self, name: Expression) -> Expression {
        // The original body reads host environment before evaluating `name`.
        // Only a literal can move to an argument without moving binding reads,
        // calls or generated-local captures across that live host boundary.
        let Some(provider) = self
            .environment
            .as_ref()
            .filter(|_| matches!(&name.kind, ExprKind::Lit(Literal::Str(_))))
        else {
            return cobol_env_lookup(name);
        };
        provider.requested.set(true);
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(&provider.name)),
            args: vec![Argument::positional(name)],
            optional: false,
        })
    }

    pub fn publish(&self, module: &mut Module) {
        fn publish(
            module: &mut Module,
            provider: &Provider,
            original: Expression,
            params: Vec<Param>,
        ) {
            let ExprKind::Call { callee, .. } = original.kind else {
                unreachable!()
            };
            let ExprKind::Lambda {
                params: original_params,
                body: LambdaBody::Block(body),
                captures,
                is_async,
            } = callee.kind
            else {
                unreachable!()
            };
            assert!(original_params.is_empty() && captures.is_empty() && !is_async);
            module.body.insert(
                0,
                Statement::new(StmtKind::FunctionDecl {
                    name: provider.name.clone(),
                    params,
                    return_type: None,
                    body,
                    modifiers: Modifiers::default(),
                    handles: Vec::new(),
                    is_async: false,
                    is_generator: false,
                    is_sub: false,
                }),
            );
        }
        if let Some(provider) = self.command_line.as_ref().filter(|p| p.requested.get()) {
            publish(module, provider, cobol_join_args_inline(), Vec::new());
        }
        if let Some(provider) = self.environment.as_ref().filter(|p| p.requested.get()) {
            let parameter = "__env_literal_name";
            publish(
                module,
                provider,
                cobol_env_lookup(Expression::ident(parameter)),
                vec![Param {
                    name: parameter.to_string(),
                    type_hint: None,
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                }],
            );
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    fn module() -> Module {
        Module {
            name: String::new(),
            language: Lang::Cobol,
            body: Vec::new(),
            imports: Vec::new(),
            directives: Directives::default(),
            canon: Default::default(),
        }
    }
    #[test]
    fn repeated_provider_sites_share_one_unchanged_body() {
        let request =
            Requests::for_source("ACCEPT A FROM COMMAND-LINE. ACCEPT B FROM COMMAND-LINE.");
        let nested = request.clone();
        let mut actual = module();
        actual
            .body
            .push(Statement::new(StmtKind::Expr(request.command_line())));
        actual
            .body
            .push(Statement::new(StmtKind::Expr(nested.command_line())));
        request.publish(&mut actual);
        assert_eq!(
            actual
                .body
                .iter()
                .filter(|s| matches!(s.kind, StmtKind::FunctionDecl { .. }))
                .count(),
            1
        );
        let StmtKind::FunctionDecl { name, body, .. } = &actual.body[0].kind else {
            unreachable!()
        };
        let ExprKind::Call { callee, .. } = cobol_join_args_inline().kind else {
            unreachable!()
        };
        let ExprKind::Lambda {
            body: LambdaBody::Block(original),
            ..
        } = callee.kind
        else {
            unreachable!()
        };
        assert_eq!(format!("{body:?}"), format!("{original:?}"));
        for statement in &actual.body[1..] {
            let StmtKind::Expr(Expression {
                kind: ExprKind::Call { callee, args, .. },
                ..
            }) = &statement.kind
            else {
                unreachable!()
            };
            assert!(matches!(&callee.kind, ExprKind::Ident(target) if target == name));
            assert!(args.is_empty());
        }
    }
    #[test]
    fn source_dependency_shadow_and_unowned_context_keep_inline() {
        for request in [
            Requests::for_source("COMMAND-LINE __COBOL_ARGS"),
            Requests::default(),
        ] {
            let ExprKind::Call { callee, .. } = request.command_line().kind else {
                unreachable!()
            };
            assert!(matches!(callee.kind, ExprKind::Lambda { .. }));
            let mut actual = module();
            request.publish(&mut actual);
            assert!(actual.body.is_empty());
        }
    }
    #[test]
    fn unused_provider_does_not_add_code_and_literal_name_is_reserved() {
        let request = Requests::for_source("COMMAND-LINE '<VYBE:COBOL.COMMAND-LINE:0>'");
        let mut actual = module();
        request.publish(&mut actual);
        assert!(actual.body.is_empty());
        let ExprKind::Call { callee, .. } = request.command_line().kind else {
            unreachable!()
        };
        assert!(
            matches!(&callee.kind, ExprKind::Ident(name) if name == "<vybe:cobol.command-line:1>")
        );
    }

    #[test]
    fn environment_literals_share_body_but_dynamic_names_keep_original_staging() {
        let request = Requests::for_source("ENVIRONMENT");
        for name in [
            Expression::ident("__env_pairs"),
            cobol_call("next_name", vec![]),
        ] {
            let actual = request.environment(name.clone());
            assert_eq!(
                format!("{actual:?}"),
                format!("{:?}", cobol_env_lookup(name))
            );
        }
        let mut actual = module();
        request.publish(&mut actual);
        assert!(actual.body.is_empty());
        for name in ["HOME", "PATH"] {
            let ExprKind::Call { callee, args, .. } =
                request.environment(Expression::string(name)).kind
            else {
                unreachable!()
            };
            assert!(
                matches!(&callee.kind, ExprKind::Ident(target) if target == "<vybe:cobol.environment-literal:0>")
            );
            assert!(
                matches!(&args[0].value.kind, ExprKind::Lit(Literal::Str(value)) if value == name)
            );
        }
        request.publish(&mut actual);
        let StmtKind::FunctionDecl { params, body, .. } = &actual.body[0].kind else {
            unreachable!()
        };
        assert_eq!(actual.body.len(), 1);
        assert_eq!(params.len(), 1);
        // Identical original statement sequence: host read, then name binding,
        // then result/loop. Only the literal is replaced by a value parameter.
        let ExprKind::Call { callee, .. } =
            cobol_env_lookup(Expression::ident(&params[0].name)).kind
        else {
            unreachable!()
        };
        let ExprKind::Lambda {
            body: LambdaBody::Block(original),
            ..
        } = callee.kind
        else {
            unreachable!()
        };
        assert_eq!(format!("{body:?}"), format!("{original:?}"));
        let shadowed = Requests::for_source("ENVIRONMENT __COBOL_ENV_PAIRS");
        let ExprKind::Call { callee, .. } = shadowed.environment(Expression::string("HOME")).kind
        else {
            unreachable!()
        };
        assert!(matches!(callee.kind, ExprKind::Lambda { .. }));
    }
}
