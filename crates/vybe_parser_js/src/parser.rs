use crate::ast::*;
use crate::lexer::Lexer;
use crate::token::{Token, TokenKind};

pub struct Parser {
    tokens: Vec<Token>,
    pos: usize,
    lexer: Lexer,
}

impl Parser {
    pub fn new(source: &str) -> Self {
        let mut lexer = Lexer::new(source);
        let tokens = lexer.tokenize().expect("Lexer error");
        Parser { tokens, pos: 0, lexer }
    }

    pub fn parse_program(&mut self) -> Result<Program, String> {
        let mut body = Vec::new();
        while !self.is_at_end() {
            body.push(self.parse_statement()?);
        }
        Ok(Program { body })
    }

    // -- Statements --

    fn parse_statement(&mut self) -> Result<Statement, String> {
        match self.current_kind() {
            TokenKind::LBrace => self.parse_block_statement(),
            TokenKind::Var | TokenKind::Let | TokenKind::Const => self.parse_variable_declaration(),
            TokenKind::Function => self.parse_function_declaration(),
            TokenKind::Class => self.parse_class_declaration(),
            TokenKind::If => self.parse_if_statement(),
            TokenKind::For => self.parse_for_statement(),
            TokenKind::While => self.parse_while_statement(),
            TokenKind::Do => self.parse_do_while_statement(),
            TokenKind::Switch => self.parse_switch_statement(),
            TokenKind::Return => self.parse_return_statement(),
            TokenKind::Break => self.parse_break_statement(),
            TokenKind::Continue => self.parse_continue_statement(),
            TokenKind::Throw => self.parse_throw_statement(),
            TokenKind::Try => self.parse_try_statement(),
            TokenKind::Import => self.parse_import_statement(),
            TokenKind::Export => self.parse_export_statement(),
            TokenKind::Semicolon => { self.advance(); Ok(Statement::Empty) }
            _ => self.parse_expression_statement(),
        }
    }

    fn parse_import_statement(&mut self) -> Result<Statement, String> {
        self.expect(TokenKind::Import)?;
        let mut specifiers = Vec::new();

        match self.current_kind() {
            // import "module" (side-effect import)
            TokenKind::String(_) => {
                let source = match self.current_kind() {
                    TokenKind::String(s) => { self.advance(); s }
                    _ => unreachable!(),
                };
                self.eat_semicolon();
                return Ok(Statement::Import { specifiers: vec![], source });
            }
            // import * as name from "module"
            TokenKind::Star => {
                self.advance();
                self.expect(TokenKind::As)?;
                let name = self.expect_identifier()?;
                specifiers.push(ImportSpecifier::Namespace(name));
            }
            // import { ... } from "module"
            TokenKind::LBrace => {
                self.advance();
                while !self.check_kind(&TokenKind::RBrace) && !self.is_at_end() {
                    let name = self.expect_identifier()?;
                    let alias = if self.eat(TokenKind::As) {
                        Some(self.expect_identifier()?)
                    } else {
                        None
                    };
                    specifiers.push(ImportSpecifier::Named { name, alias });
                    if !self.eat(TokenKind::Comma) { break; }
                }
                self.expect(TokenKind::RBrace)?;
            }
            // import defaultName from "module" or import defaultName, { ... } from "module"
            _ => {
                let name = self.expect_identifier()?;
                specifiers.push(ImportSpecifier::Default(name));
                if self.eat(TokenKind::Comma) {
                    // import default, { named } from "module"
                    if self.eat(TokenKind::LBrace) {
                        while !self.check_kind(&TokenKind::RBrace) && !self.is_at_end() {
                            let n = self.expect_identifier()?;
                            let a = if self.eat(TokenKind::As) {
                                Some(self.expect_identifier()?)
                            } else {
                                None
                            };
                            specifiers.push(ImportSpecifier::Named { name: n, alias: a });
                            if !self.eat(TokenKind::Comma) { break; }
                        }
                        self.expect(TokenKind::RBrace)?;
                    }
                }
            }
        }

        // "from" keyword + source string
        self.expect(TokenKind::From)?;
        let source = match self.current_kind() {
            TokenKind::String(s) => { self.advance(); s }
            _ => return Err(self.error("Expected module path string")),
        };
        self.eat_semicolon();
        Ok(Statement::Import { specifiers, source })
    }

    fn parse_export_statement(&mut self) -> Result<Statement, String> {
        self.expect(TokenKind::Export)?;

        // export default expr
        if self.eat(TokenKind::Default) {
            if self.check_kind(&TokenKind::Function) {
                let stmt = self.parse_function_declaration()?;
                return Ok(Statement::Export {
                    declaration: Some(Box::new(stmt)),
                    specifiers: vec![],
                    default: None,
                });
            }
            if self.check_kind(&TokenKind::Class) {
                let stmt = self.parse_class_declaration()?;
                return Ok(Statement::Export {
                    declaration: Some(Box::new(stmt)),
                    specifiers: vec![],
                    default: None,
                });
            }
            let expr = self.parse_assignment_expression()?;
            self.eat_semicolon();
            return Ok(Statement::Export {
                declaration: None,
                specifiers: vec![],
                default: Some(Box::new(expr)),
            });
        }

        // export { a, b, c }
        if self.check_kind(&TokenKind::LBrace) {
            self.advance();
            let mut specifiers = Vec::new();
            while !self.check_kind(&TokenKind::RBrace) && !self.is_at_end() {
                let name = self.expect_identifier()?;
                let alias = if self.eat(TokenKind::As) {
                    Some(self.expect_identifier()?)
                } else {
                    None
                };
                specifiers.push(ExportSpecifier { name, alias });
                if !self.eat(TokenKind::Comma) { break; }
            }
            self.expect(TokenKind::RBrace)?;
            self.eat_semicolon();
            return Ok(Statement::Export {
                declaration: None,
                specifiers,
                default: None,
            });
        }

        // export function/class/let/const/var
        let stmt = match self.current_kind() {
            TokenKind::Function => self.parse_function_declaration()?,
            TokenKind::Class => self.parse_class_declaration()?,
            TokenKind::Var | TokenKind::Let | TokenKind::Const => self.parse_variable_declaration()?,
            _ => return Err(self.error("Expected declaration after export")),
        };

        Ok(Statement::Export {
            declaration: Some(Box::new(stmt)),
            specifiers: vec![],
            default: None,
        })
    }

    fn parse_block_statement(&mut self) -> Result<Statement, String> {
        self.expect(TokenKind::LBrace)?;
        let mut stmts = Vec::new();
        while !self.check_kind(&TokenKind::RBrace) && !self.is_at_end() {
            stmts.push(self.parse_statement()?);
        }
        self.expect(TokenKind::RBrace)?;
        Ok(Statement::Block(stmts))
    }

    fn parse_block_body(&mut self) -> Result<Vec<Statement>, String> {
        self.expect(TokenKind::LBrace)?;
        let mut stmts = Vec::new();
        while !self.check_kind(&TokenKind::RBrace) && !self.is_at_end() {
            stmts.push(self.parse_statement()?);
        }
        self.expect(TokenKind::RBrace)?;
        Ok(stmts)
    }

    fn parse_variable_declaration(&mut self) -> Result<Statement, String> {
        let kind = match self.current_kind() {
            TokenKind::Var => VarKind::Var,
            TokenKind::Let => VarKind::Let,
            TokenKind::Const => VarKind::Const,
            _ => return Err(self.error("Expected var, let, or const")),
        };
        self.advance();

        let mut declarations = Vec::new();
        loop {
            let pattern = self.parse_binding_pattern()?;
            let init = if self.eat(TokenKind::Eq) {
                Some(self.parse_assignment_expression()?)
            } else {
                None
            };
            declarations.push(VarDeclarator { pattern, init });
            if !self.eat(TokenKind::Comma) {
                break;
            }
        }
        self.eat_semicolon();
        Ok(Statement::VariableDeclaration { kind, declarations })
    }

    fn parse_binding_pattern(&mut self) -> Result<BindingPattern, String> {
        match self.current_kind() {
            TokenKind::LBrace => self.parse_object_pattern(),
            TokenKind::LBracket => self.parse_array_pattern(),
            _ => {
                let name = self.expect_identifier()?;
                Ok(BindingPattern::Identifier(name))
            }
        }
    }

    fn parse_object_pattern(&mut self) -> Result<BindingPattern, String> {
        self.expect(TokenKind::LBrace)?;
        let mut props = Vec::new();
        while !self.check_kind(&TokenKind::RBrace) && !self.is_at_end() {
            if self.check_kind(&TokenKind::DotDotDot) {
                self.advance();
                let name = self.expect_identifier()?;
                props.push(ObjectPatternProp { key: name.clone(), value: Some(BindingPattern::Identifier(name)), default: None });
            } else {
                let key = self.expect_identifier()?;
                if self.eat(TokenKind::Colon) {
                    // { key: pattern }
                    let pat = self.parse_binding_pattern()?;
                    let default = if self.eat(TokenKind::Eq) { Some(self.parse_assignment_expression()?) } else { None };
                    props.push(ObjectPatternProp { key, value: Some(pat), default });
                } else {
                    // Shorthand { key } or { key = default }
                    let default = if self.eat(TokenKind::Eq) { Some(self.parse_assignment_expression()?) } else { None };
                    props.push(ObjectPatternProp { key, value: None, default });
                }
            }
            if !self.eat(TokenKind::Comma) { break; }
        }
        self.expect(TokenKind::RBrace)?;
        Ok(BindingPattern::Object(props))
    }

    fn parse_array_pattern(&mut self) -> Result<BindingPattern, String> {
        self.expect(TokenKind::LBracket)?;
        let mut elems = Vec::new();
        while !self.check_kind(&TokenKind::RBracket) && !self.is_at_end() {
            if self.check_kind(&TokenKind::Comma) {
                elems.push(ArrayPatternElem::Hole);
            } else if self.check_kind(&TokenKind::DotDotDot) {
                self.advance();
                let name = self.expect_identifier()?;
                elems.push(ArrayPatternElem::Rest(name));
                break; // rest must be last
            } else {
                let pat = self.parse_binding_pattern()?;
                let default = if self.eat(TokenKind::Eq) { Some(self.parse_assignment_expression()?) } else { None };
                elems.push(ArrayPatternElem::Pattern(pat, default));
            }
            if !self.eat(TokenKind::Comma) { break; }
        }
        self.expect(TokenKind::RBracket)?;
        Ok(BindingPattern::Array(elems))
    }

    fn parse_function_declaration(&mut self) -> Result<Statement, String> {
        let func = self.parse_function(true)?;
        Ok(Statement::FunctionDeclaration(func))
    }

    fn parse_function(&mut self, require_name: bool) -> Result<FunctionDecl, String> {
        self.expect(TokenKind::Function)?;
        let name = if require_name || self.is_identifier() {
            Some(self.expect_identifier()?)
        } else {
            None
        };
        self.expect(TokenKind::LParen)?;
        let params = self.parse_param_list()?;
        self.expect(TokenKind::RParen)?;
        let body = self.parse_block_body()?;
        Ok(FunctionDecl { name, params, body })
    }

    fn parse_param_list(&mut self) -> Result<Vec<Param>, String> {
        let mut params = Vec::new();
        if !self.check_kind(&TokenKind::RParen) {
            loop {
                let rest = if self.check_kind(&TokenKind::DotDotDot) {
                    self.advance();
                    true
                } else {
                    false
                };
                let name = self.expect_identifier()?;
                let default = if self.eat(TokenKind::Eq) {
                    Some(self.parse_assignment_expression()?)
                } else {
                    None
                };
                params.push(Param { name, default, rest });
                if rest || !self.eat(TokenKind::Comma) {
                    break;
                }
            }
        }
        Ok(params)
    }

    fn parse_class_declaration(&mut self) -> Result<Statement, String> {
        let class = self.parse_class()?;
        Ok(Statement::ClassDeclaration(class))
    }

    fn parse_class(&mut self) -> Result<ClassDecl, String> {
        self.expect(TokenKind::Class)?;
        let name = if self.is_identifier() {
            Some(self.expect_identifier()?)
        } else {
            None
        };
        let super_class = if self.eat(TokenKind::Extends) {
            Some(Box::new(self.parse_assignment_expression()?))
        } else {
            None
        };

        self.expect(TokenKind::LBrace)?;
        let mut body = Vec::new();
        while !self.check_kind(&TokenKind::RBrace) && !self.is_at_end() {
            if self.check_kind(&TokenKind::Semicolon) {
                self.advance();
                continue;
            }
            body.push(self.parse_class_member()?);
        }
        self.expect(TokenKind::RBrace)?;

        Ok(ClassDecl { name, super_class, body })
    }

    fn parse_class_member(&mut self) -> Result<ClassMember, String> {
        let is_static = self.eat(TokenKind::Static);

        // Check for getter/setter
        let kind = if self.check_kind(&TokenKind::Get) && self.peek_is_method_name() {
            self.advance();
            MethodKind::Get
        } else if self.check_kind(&TokenKind::Set) && self.peek_is_method_name() {
            self.advance();
            MethodKind::Set
        } else {
            MethodKind::Method
        };

        let key = self.expect_property_name()?;

        let actual_kind = if key == "constructor" && !is_static && kind == MethodKind::Method {
            MethodKind::Constructor
        } else {
            kind
        };

        if self.check_kind(&TokenKind::LParen) {
            // Method
            self.expect(TokenKind::LParen)?;
            let params = self.parse_param_list()?;
            self.expect(TokenKind::RParen)?;
            let body = self.parse_block_body()?;
            Ok(ClassMember::Method {
                key,
                value: FunctionDecl { name: None, params, body },
                kind: actual_kind,
                is_static,
            })
        } else {
            // Property
            let value = if self.eat(TokenKind::Eq) {
                Some(self.parse_assignment_expression()?)
            } else {
                None
            };
            self.eat_semicolon();
            Ok(ClassMember::Property { key, value, is_static })
        }
    }

    fn parse_if_statement(&mut self) -> Result<Statement, String> {
        self.expect(TokenKind::If)?;
        self.expect(TokenKind::LParen)?;
        let test = self.parse_expression()?;
        self.expect(TokenKind::RParen)?;
        let consequent = Box::new(self.parse_statement()?);
        let alternate = if self.eat(TokenKind::Else) {
            Some(Box::new(self.parse_statement()?))
        } else {
            None
        };
        Ok(Statement::If { test, consequent, alternate })
    }

    fn parse_for_statement(&mut self) -> Result<Statement, String> {
        self.expect(TokenKind::For)?;
        self.expect(TokenKind::LParen)?;

        // Check for for-in / for-of
        if self.check_kind(&TokenKind::Var) || self.check_kind(&TokenKind::Let) || self.check_kind(&TokenKind::Const) {
            let kind = match self.current_kind() {
                TokenKind::Var => VarKind::Var,
                TokenKind::Let => VarKind::Let,
                TokenKind::Const => VarKind::Const,
                _ => unreachable!(),
            };
            self.advance();
            let name = self.expect_identifier()?;

            if self.eat(TokenKind::In) {
                let right = self.parse_expression()?;
                self.expect(TokenKind::RParen)?;
                let body = Box::new(self.parse_statement()?);
                return Ok(Statement::ForIn {
                    left: ForInTarget::VarDecl(kind, name),
                    right, body,
                });
            }
            if self.eat(TokenKind::Of) {
                let right = self.parse_expression()?;
                self.expect(TokenKind::RParen)?;
                let body = Box::new(self.parse_statement()?);
                return Ok(Statement::ForOf {
                    left: ForInTarget::VarDecl(kind, name),
                    right, body,
                });
            }

            // Regular for with var decl init
            let mut declarations = vec![];
            let init_expr = if self.eat(TokenKind::Eq) {
                Some(self.parse_assignment_expression()?)
            } else {
                None
            };
            declarations.push(VarDeclarator::simple(name, init_expr));
            while self.eat(TokenKind::Comma) {
                let n = self.expect_identifier()?;
                let init = if self.eat(TokenKind::Eq) {
                    Some(self.parse_assignment_expression()?)
                } else {
                    None
                };
                declarations.push(VarDeclarator::simple(n, init));
            }

            self.expect(TokenKind::Semicolon)?;
            let test = if !self.check_kind(&TokenKind::Semicolon) {
                Some(self.parse_expression()?)
            } else { None };
            self.expect(TokenKind::Semicolon)?;
            let update = if !self.check_kind(&TokenKind::RParen) {
                Some(self.parse_expression()?)
            } else { None };
            self.expect(TokenKind::RParen)?;
            let body = Box::new(self.parse_statement()?);

            return Ok(Statement::For {
                init: Some(ForInit::VarDecl(kind, declarations)),
                test, update, body,
            });
        }

        // Regular for or for-in/of with expression init
        let init = if self.check_kind(&TokenKind::Semicolon) {
            None
        } else {
            Some(ForInit::Expression(self.parse_expression()?))
        };
        self.expect(TokenKind::Semicolon)?;
        let test = if !self.check_kind(&TokenKind::Semicolon) {
            Some(self.parse_expression()?)
        } else { None };
        self.expect(TokenKind::Semicolon)?;
        let update = if !self.check_kind(&TokenKind::RParen) {
            Some(self.parse_expression()?)
        } else { None };
        self.expect(TokenKind::RParen)?;
        let body = Box::new(self.parse_statement()?);

        Ok(Statement::For { init, test, update, body })
    }

    fn parse_while_statement(&mut self) -> Result<Statement, String> {
        self.expect(TokenKind::While)?;
        self.expect(TokenKind::LParen)?;
        let test = self.parse_expression()?;
        self.expect(TokenKind::RParen)?;
        let body = Box::new(self.parse_statement()?);
        Ok(Statement::While { test, body })
    }

    fn parse_do_while_statement(&mut self) -> Result<Statement, String> {
        self.expect(TokenKind::Do)?;
        let body = Box::new(self.parse_statement()?);
        self.expect(TokenKind::While)?;
        self.expect(TokenKind::LParen)?;
        let test = self.parse_expression()?;
        self.expect(TokenKind::RParen)?;
        self.eat_semicolon();
        Ok(Statement::DoWhile { body, test })
    }

    fn parse_switch_statement(&mut self) -> Result<Statement, String> {
        self.expect(TokenKind::Switch)?;
        self.expect(TokenKind::LParen)?;
        let discriminant = self.parse_expression()?;
        self.expect(TokenKind::RParen)?;
        self.expect(TokenKind::LBrace)?;

        let mut cases = Vec::new();
        while !self.check_kind(&TokenKind::RBrace) && !self.is_at_end() {
            let test = if self.eat(TokenKind::Case) {
                Some(self.parse_expression()?)
            } else {
                self.expect(TokenKind::Default)?;
                None
            };
            self.expect(TokenKind::Colon)?;
            let mut consequent = Vec::new();
            while !self.check_kind(&TokenKind::Case) && !self.check_kind(&TokenKind::Default)
                && !self.check_kind(&TokenKind::RBrace) && !self.is_at_end()
            {
                consequent.push(self.parse_statement()?);
            }
            cases.push(SwitchCase { test, consequent });
        }
        self.expect(TokenKind::RBrace)?;
        Ok(Statement::Switch { discriminant, cases })
    }

    fn parse_return_statement(&mut self) -> Result<Statement, String> {
        self.expect(TokenKind::Return)?;
        let value = if self.check_kind(&TokenKind::Semicolon) || self.check_kind(&TokenKind::RBrace) || self.is_at_end() {
            None
        } else {
            Some(self.parse_expression()?)
        };
        self.eat_semicolon();
        Ok(Statement::Return(value))
    }

    fn parse_break_statement(&mut self) -> Result<Statement, String> {
        self.expect(TokenKind::Break)?;
        let label = if self.is_identifier() && !self.prev_followed_by_newline() {
            Some(self.expect_identifier()?)
        } else { None };
        self.eat_semicolon();
        Ok(Statement::Break(label))
    }

    fn parse_continue_statement(&mut self) -> Result<Statement, String> {
        self.expect(TokenKind::Continue)?;
        let label = if self.is_identifier() && !self.prev_followed_by_newline() {
            Some(self.expect_identifier()?)
        } else { None };
        self.eat_semicolon();
        Ok(Statement::Continue(label))
    }

    fn parse_throw_statement(&mut self) -> Result<Statement, String> {
        self.expect(TokenKind::Throw)?;
        let expr = self.parse_expression()?;
        self.eat_semicolon();
        Ok(Statement::Throw(expr))
    }

    fn parse_try_statement(&mut self) -> Result<Statement, String> {
        self.expect(TokenKind::Try)?;
        let block = self.parse_block_body()?;
        let handler = if self.eat(TokenKind::Catch) {
            let param = if self.eat(TokenKind::LParen) {
                let p = self.expect_identifier()?;
                self.expect(TokenKind::RParen)?;
                Some(p)
            } else { None };
            let body = self.parse_block_body()?;
            Some(CatchClause { param, body })
        } else { None };
        let finalizer = if self.eat(TokenKind::Finally) {
            Some(self.parse_block_body()?)
        } else { None };
        Ok(Statement::Try { block, handler, finalizer })
    }

    fn parse_expression_statement(&mut self) -> Result<Statement, String> {
        let expr = self.parse_expression()?;
        self.eat_semicolon();
        Ok(Statement::Expression(expr))
    }

    // -- Expressions (Pratt / operator precedence) --

    fn parse_expression(&mut self) -> Result<Expression, String> {
        let expr = self.parse_assignment_expression()?;
        // Handle comma expressions
        if self.check_kind(&TokenKind::Comma) && !self.is_in_call_args() {
            let mut exprs = vec![expr];
            while self.eat(TokenKind::Comma) {
                exprs.push(self.parse_assignment_expression()?);
            }
            Ok(Expression::Sequence(exprs))
        } else {
            Ok(expr)
        }
    }

    fn parse_assignment_expression(&mut self) -> Result<Expression, String> {
        let left = self.parse_conditional()?;

        if self.current_kind().is_assignment_op() {
            let op = self.parse_assign_op()?;
            let right = self.parse_assignment_expression()?;
            return Ok(Expression::Assignment {
                op,
                left: Box::new(left),
                right: Box::new(right),
            });
        }

        Ok(left)
    }

    fn parse_conditional(&mut self) -> Result<Expression, String> {
        let mut expr = self.parse_nullish_coalescing()?;
        if self.eat(TokenKind::Question) {
            let consequent = self.parse_assignment_expression()?;
            self.expect(TokenKind::Colon)?;
            let alternate = self.parse_assignment_expression()?;
            expr = Expression::Conditional {
                test: Box::new(expr),
                consequent: Box::new(consequent),
                alternate: Box::new(alternate),
            };
        }
        Ok(expr)
    }

    fn parse_nullish_coalescing(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_logical_or()?;
        while self.eat(TokenKind::QuestionQuestion) {
            let right = self.parse_logical_or()?;
            left = Expression::Binary {
                op: BinaryOp::NullishCoalescing,
                left: Box::new(left),
                right: Box::new(right),
            };
        }
        Ok(left)
    }

    fn parse_logical_or(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_logical_and()?;
        while self.eat(TokenKind::PipePipe) {
            let right = self.parse_logical_and()?;
            left = Expression::Logical {
                op: LogicalOp::Or,
                left: Box::new(left),
                right: Box::new(right),
            };
        }
        Ok(left)
    }

    fn parse_logical_and(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_bitwise_or()?;
        while self.eat(TokenKind::AmpAmp) {
            let right = self.parse_bitwise_or()?;
            left = Expression::Logical {
                op: LogicalOp::And,
                left: Box::new(left),
                right: Box::new(right),
            };
        }
        Ok(left)
    }

    fn parse_bitwise_or(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_bitwise_xor()?;
        while self.eat(TokenKind::Pipe) {
            let right = self.parse_bitwise_xor()?;
            left = Expression::Binary { op: BinaryOp::BitOr, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_bitwise_xor(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_bitwise_and()?;
        while self.eat(TokenKind::Caret) {
            let right = self.parse_bitwise_and()?;
            left = Expression::Binary { op: BinaryOp::BitXor, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_bitwise_and(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_equality()?;
        while self.eat(TokenKind::Amp) {
            let right = self.parse_equality()?;
            left = Expression::Binary { op: BinaryOp::BitAnd, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_equality(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_relational()?;
        loop {
            let op = match self.current_kind() {
                TokenKind::EqEq => BinaryOp::Eq,
                TokenKind::BangEq => BinaryOp::Neq,
                TokenKind::EqEqEq => BinaryOp::SEq,
                TokenKind::BangEqEq => BinaryOp::SNeq,
                _ => break,
            };
            self.advance();
            let right = self.parse_relational()?;
            left = Expression::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_relational(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_shift()?;
        loop {
            let op = match self.current_kind() {
                TokenKind::Lt => BinaryOp::Lt,
                TokenKind::Gt => BinaryOp::Gt,
                TokenKind::LtEq => BinaryOp::Le,
                TokenKind::GtEq => BinaryOp::Ge,
                TokenKind::Instanceof => BinaryOp::InstanceOf,
                TokenKind::In => BinaryOp::In,
                _ => break,
            };
            self.advance();
            let right = self.parse_shift()?;
            left = Expression::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_shift(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_additive()?;
        loop {
            let op = match self.current_kind() {
                TokenKind::LtLt => BinaryOp::Shl,
                TokenKind::GtGt => BinaryOp::Shr,
                TokenKind::GtGtGt => BinaryOp::UShr,
                _ => break,
            };
            self.advance();
            let right = self.parse_additive()?;
            left = Expression::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_additive(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_multiplicative()?;
        loop {
            let op = match self.current_kind() {
                TokenKind::Plus => BinaryOp::Add,
                TokenKind::Minus => BinaryOp::Sub,
                _ => break,
            };
            self.advance();
            let right = self.parse_multiplicative()?;
            left = Expression::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_multiplicative(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_exponentiation()?;
        loop {
            let op = match self.current_kind() {
                TokenKind::Star => BinaryOp::Mul,
                TokenKind::Slash => BinaryOp::Div,
                TokenKind::Percent => BinaryOp::Mod,
                _ => break,
            };
            self.advance();
            let right = self.parse_exponentiation()?;
            left = Expression::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_exponentiation(&mut self) -> Result<Expression, String> {
        let left = self.parse_unary()?;
        if self.eat(TokenKind::StarStar) {
            let right = self.parse_exponentiation()?; // right-associative
            Ok(Expression::Binary { op: BinaryOp::Exp, left: Box::new(left), right: Box::new(right) })
        } else {
            Ok(left)
        }
    }

    fn parse_unary(&mut self) -> Result<Expression, String> {
        match self.current_kind() {
            TokenKind::Minus => {
                self.advance();
                let arg = self.parse_unary()?;
                Ok(Expression::Unary { op: UnaryOp::Neg, argument: Box::new(arg) })
            }
            TokenKind::Plus => {
                self.advance();
                let arg = self.parse_unary()?;
                Ok(Expression::Unary { op: UnaryOp::Pos, argument: Box::new(arg) })
            }
            TokenKind::Bang => {
                self.advance();
                let arg = self.parse_unary()?;
                Ok(Expression::Unary { op: UnaryOp::Not, argument: Box::new(arg) })
            }
            TokenKind::Tilde => {
                self.advance();
                let arg = self.parse_unary()?;
                Ok(Expression::Unary { op: UnaryOp::BitNot, argument: Box::new(arg) })
            }
            TokenKind::Typeof => {
                self.advance();
                let arg = self.parse_unary()?;
                Ok(Expression::Typeof(Box::new(arg)))
            }
            TokenKind::Void => {
                self.advance();
                let arg = self.parse_unary()?;
                Ok(Expression::Void(Box::new(arg)))
            }
            TokenKind::Delete => {
                self.advance();
                let arg = self.parse_unary()?;
                Ok(Expression::Delete(Box::new(arg)))
            }
            TokenKind::PlusPlus => {
                self.advance();
                let arg = self.parse_unary()?;
                Ok(Expression::Update { op: UpdateOp::Increment, prefix: true, argument: Box::new(arg) })
            }
            TokenKind::MinusMinus => {
                self.advance();
                let arg = self.parse_unary()?;
                Ok(Expression::Update { op: UpdateOp::Decrement, prefix: true, argument: Box::new(arg) })
            }
            _ => self.parse_postfix(),
        }
    }

    fn parse_postfix(&mut self) -> Result<Expression, String> {
        let mut expr = self.parse_call_expression()?;

        // Postfix ++ and --
        if !self.prev_followed_by_newline() {
            if self.check_kind(&TokenKind::PlusPlus) {
                self.advance();
                expr = Expression::Update { op: UpdateOp::Increment, prefix: false, argument: Box::new(expr) };
            } else if self.check_kind(&TokenKind::MinusMinus) {
                self.advance();
                expr = Expression::Update { op: UpdateOp::Decrement, prefix: false, argument: Box::new(expr) };
            }
        }

        Ok(expr)
    }

    fn parse_call_expression(&mut self) -> Result<Expression, String> {
        let mut expr = if self.check_kind(&TokenKind::New) {
            self.parse_new_expression()?
        } else {
            self.parse_primary()?
        };

        loop {
            match self.current_kind() {
                TokenKind::LParen => {
                    self.advance();
                    let args = self.parse_arguments()?;
                    self.expect(TokenKind::RParen)?;
                    expr = Expression::Call { callee: Box::new(expr), arguments: args, optional: false };
                }
                TokenKind::Dot => {
                    self.advance();
                    let prop = self.expect_property_name()?;
                    expr = Expression::Member { object: Box::new(expr), property: prop, optional: false };
                }
                TokenKind::QuestionDot => {
                    self.advance();
                    if self.check_kind(&TokenKind::LParen) {
                        self.advance();
                        let args = self.parse_arguments()?;
                        self.expect(TokenKind::RParen)?;
                        expr = Expression::Call { callee: Box::new(expr), arguments: args, optional: true };
                    } else {
                        let prop = self.expect_property_name()?;
                        expr = Expression::Member { object: Box::new(expr), property: prop, optional: true };
                    }
                }
                TokenKind::LBracket => {
                    self.advance();
                    let prop = self.parse_expression()?;
                    self.expect(TokenKind::RBracket)?;
                    expr = Expression::ComputedMember { object: Box::new(expr), property: Box::new(prop) };
                }
                _ => break,
            }
        }

        Ok(expr)
    }

    fn parse_new_expression(&mut self) -> Result<Expression, String> {
        self.expect(TokenKind::New)?;
        let callee = self.parse_call_expression()?;
        // If callee was already parsed as a Call, convert it
        match callee {
            Expression::Call { callee, arguments, .. } => {
                Ok(Expression::New { callee, arguments })
            }
            _ => {
                Ok(Expression::New { callee: Box::new(callee), arguments: vec![] })
            }
        }
    }

    fn parse_arguments(&mut self) -> Result<Vec<Expression>, String> {
        let mut args = Vec::new();
        if !self.check_kind(&TokenKind::RParen) {
            loop {
                if self.check_kind(&TokenKind::DotDotDot) {
                    self.advance();
                    let expr = self.parse_assignment_expression()?;
                    args.push(Expression::Spread(Box::new(expr)));
                } else {
                    args.push(self.parse_assignment_expression()?);
                }
                if !self.eat(TokenKind::Comma) {
                    break;
                }
                // Trailing comma
                if self.check_kind(&TokenKind::RParen) {
                    break;
                }
            }
        }
        Ok(args)
    }

    fn parse_primary(&mut self) -> Result<Expression, String> {
        match self.current_kind() {
            TokenKind::Number(n) => {
                let val = n;
                self.advance();
                Ok(Expression::Number(val))
            }
            TokenKind::String(ref s) => {
                let val = s.clone();
                self.advance();
                Ok(Expression::String(val))
            }
            TokenKind::TemplateLiteral(ref s) => {
                let val = s.clone();
                self.advance();
                Ok(Expression::TemplateLiteral { quasis: vec![val], expressions: vec![] })
            }
            TokenKind::TemplateHead(ref s) => {
                let head = s.clone();
                self.advance();
                self.parse_template_expression(head)
            }
            TokenKind::True => { self.advance(); Ok(Expression::Boolean(true)) }
            TokenKind::False => { self.advance(); Ok(Expression::Boolean(false)) }
            TokenKind::Null => { self.advance(); Ok(Expression::Null) }
            TokenKind::Undefined => { self.advance(); Ok(Expression::Undefined) }
            TokenKind::This => { self.advance(); Ok(Expression::This) }
            TokenKind::Identifier(_) => {
                let name = self.expect_identifier()?;

                // Check for arrow function: ident => ...
                if self.check_kind(&TokenKind::Arrow) {
                    self.advance(); // skip =>
                    return self.parse_arrow_body(vec![Param::simple(name)]);
                }

                Ok(Expression::Identifier(name))
            }
            TokenKind::LParen => {
                // Could be grouped expression or arrow function params
                self.advance();

                // Empty parens = arrow function
                if self.check_kind(&TokenKind::RParen) {
                    self.advance();
                    self.expect(TokenKind::Arrow)?;
                    return self.parse_arrow_body(vec![]);
                }

                // Try to parse as arrow function params
                if self.could_be_arrow_params() {
                    if let Some(params) = self.try_parse_arrow_params()? {
                        self.expect(TokenKind::Arrow)?;
                        return self.parse_arrow_body(params);
                    }
                }

                // Regular grouped expression
                let expr = self.parse_expression()?;
                self.expect(TokenKind::RParen)?;

                // Check for arrow after parens
                if self.check_kind(&TokenKind::Arrow) {
                    if let Some(params) = self.expr_to_params(&expr) {
                        self.advance(); // skip =>
                        return self.parse_arrow_body(params);
                    }
                }

                Ok(expr)
            }
            TokenKind::LBracket => self.parse_array_literal(),
            TokenKind::LBrace => self.parse_object_literal(),
            TokenKind::Function => {
                let func = self.parse_function(false)?;
                Ok(Expression::Function(func))
            }
            TokenKind::Class => {
                let class = self.parse_class()?;
                Ok(Expression::Function(FunctionDecl {
                    name: class.name,
                    params: vec![],
                    body: vec![],
                }))
            }
            TokenKind::DotDotDot => {
                self.advance();
                let expr = self.parse_assignment_expression()?;
                Ok(Expression::Spread(Box::new(expr)))
            }
            _ => Err(self.error(&format!("Unexpected token: {:?}", self.current_kind()))),
        }
    }

    fn parse_template_expression(&mut self, head: String) -> Result<Expression, String> {
        let mut quasis = vec![head];
        let mut expressions = Vec::new();

        loop {
            expressions.push(self.parse_expression()?);

            // The lexer produces TemplateTail or TemplateMiddle after the }
            match self.current_kind() {
                TokenKind::TemplateTail(s) => {
                    quasis.push(s);
                    self.advance();
                    break;
                }
                TokenKind::TemplateMiddle(s) => {
                    quasis.push(s);
                    self.advance();
                    // Continue parsing next interpolation expression
                }
                _ => return Err(self.error("Expected template continuation")),
            }
        }

        Ok(Expression::TemplateLiteral { quasis, expressions })
    }

    fn parse_array_literal(&mut self) -> Result<Expression, String> {
        self.expect(TokenKind::LBracket)?;
        let mut elements = Vec::new();
        while !self.check_kind(&TokenKind::RBracket) && !self.is_at_end() {
            if self.check_kind(&TokenKind::DotDotDot) {
                self.advance();
                elements.push(Expression::Spread(Box::new(self.parse_assignment_expression()?)));
            } else {
                elements.push(self.parse_assignment_expression()?);
            }
            if !self.eat(TokenKind::Comma) {
                break;
            }
        }
        self.expect(TokenKind::RBracket)?;
        Ok(Expression::Array(elements))
    }

    fn parse_object_literal(&mut self) -> Result<Expression, String> {
        self.expect(TokenKind::LBrace)?;
        let mut properties = Vec::new();
        while !self.check_kind(&TokenKind::RBrace) && !self.is_at_end() {
            if self.check_kind(&TokenKind::DotDotDot) {
                self.advance();
                properties.push(PropertyDef::Spread(self.parse_assignment_expression()?));
            } else {
                let key = self.expect_property_name()?;

                if self.check_kind(&TokenKind::LParen) {
                    // Method shorthand: { foo() { ... } }
                    self.expect(TokenKind::LParen)?;
                    let params = self.parse_param_list()?;
                    self.expect(TokenKind::RParen)?;
                    let body = self.parse_block_body()?;
                    properties.push(PropertyDef::Method {
                        key,
                        value: FunctionDecl { name: None, params, body },
                    });
                } else if self.eat(TokenKind::Colon) {
                    // key: value
                    let value = self.parse_assignment_expression()?;
                    properties.push(PropertyDef::KeyValue { key, value });
                } else {
                    // Shorthand { x }
                    properties.push(PropertyDef::Shorthand(key));
                }
            }
            if !self.eat(TokenKind::Comma) {
                break;
            }
        }
        self.expect(TokenKind::RBrace)?;
        Ok(Expression::Object(properties))
    }

    fn parse_arrow_body(&mut self, params: Vec<Param>) -> Result<Expression, String> {
        let body = if self.check_kind(&TokenKind::LBrace) {
            ArrowBody::Block(self.parse_block_body()?)
        } else {
            ArrowBody::Expression(Box::new(self.parse_assignment_expression()?))
        };
        Ok(Expression::ArrowFunction { params, body })
    }

    // -- Helpers --

    fn parse_assign_op(&mut self) -> Result<AssignOp, String> {
        let op = match self.current_kind() {
            TokenKind::Eq => AssignOp::Assign,
            TokenKind::PlusEq => AssignOp::AddAssign,
            TokenKind::MinusEq => AssignOp::SubAssign,
            TokenKind::StarEq => AssignOp::MulAssign,
            TokenKind::SlashEq => AssignOp::DivAssign,
            TokenKind::PercentEq => AssignOp::ModAssign,
            TokenKind::AmpEq => AssignOp::BitAndAssign,
            TokenKind::PipeEq => AssignOp::BitOrAssign,
            TokenKind::CaretEq => AssignOp::BitXorAssign,
            TokenKind::LtLtEq => AssignOp::ShlAssign,
            TokenKind::GtGtEq => AssignOp::ShrAssign,
            TokenKind::GtGtGtEq => AssignOp::UShrAssign,
            TokenKind::StarStarEq => AssignOp::ExpAssign,
            TokenKind::AmpAmpEq => AssignOp::AndAssign,
            TokenKind::PipePipeEq => AssignOp::OrAssign,
            TokenKind::QuestionQuestionEq => AssignOp::NullishAssign,
            _ => return Err(self.error("Expected assignment operator")),
        };
        self.advance();
        Ok(op)
    }

    fn current_kind(&self) -> TokenKind {
        if self.pos < self.tokens.len() {
            self.tokens[self.pos].kind.clone()
        } else {
            TokenKind::Eof
        }
    }

    fn advance(&mut self) {
        if self.pos < self.tokens.len() {
            self.pos += 1;
        }
    }

    fn is_at_end(&self) -> bool {
        self.pos >= self.tokens.len() || self.tokens[self.pos].kind == TokenKind::Eof
    }

    fn check_kind(&self, kind: &TokenKind) -> bool {
        if self.pos < self.tokens.len() {
            std::mem::discriminant(&self.tokens[self.pos].kind) == std::mem::discriminant(kind)
        } else {
            false
        }
    }

    fn eat(&mut self, kind: TokenKind) -> bool {
        if self.check_kind(&kind) {
            self.advance();
            true
        } else {
            false
        }
    }

    fn expect(&mut self, kind: TokenKind) -> Result<(), String> {
        if self.check_kind(&kind) {
            self.advance();
            Ok(())
        } else {
            Err(self.error(&format!("Expected {:?}, got {:?}", kind, self.current_kind())))
        }
    }

    fn expect_identifier(&mut self) -> Result<String, String> {
        match self.current_kind() {
            TokenKind::Identifier(name) => {
                self.advance();
                Ok(name)
            }
            // Allow contextual keywords as identifiers
            TokenKind::Get => { self.advance(); Ok("get".to_string()) }
            TokenKind::Set => { self.advance(); Ok("set".to_string()) }
            TokenKind::From => { self.advance(); Ok("from".to_string()) }
            TokenKind::Of => { self.advance(); Ok("of".to_string()) }
            TokenKind::As => { self.advance(); Ok("as".to_string()) }
            TokenKind::Async => { self.advance(); Ok("async".to_string()) }
            _ => Err(self.error(&format!("Expected identifier, got {:?}", self.current_kind()))),
        }
    }

    fn expect_property_name(&mut self) -> Result<String, String> {
        match self.current_kind() {
            TokenKind::Identifier(name) => { self.advance(); Ok(name) }
            TokenKind::String(s) => { self.advance(); Ok(s) }
            TokenKind::Number(n) => { self.advance(); Ok(format!("{}", n)) }
            // Keywords are valid property names
            _ if self.is_keyword() => {
                let name = format!("{:?}", self.current_kind()).to_lowercase();
                self.advance();
                Ok(name)
            }
            _ => Err(self.error(&format!("Expected property name, got {:?}", self.current_kind()))),
        }
    }

    fn is_identifier(&self) -> bool {
        matches!(self.current_kind(), TokenKind::Identifier(_))
    }

    fn is_keyword(&self) -> bool {
        matches!(self.current_kind(),
            TokenKind::Var | TokenKind::Let | TokenKind::Const | TokenKind::Function |
            TokenKind::Return | TokenKind::If | TokenKind::Else | TokenKind::For |
            TokenKind::While | TokenKind::Do | TokenKind::Break | TokenKind::Continue |
            TokenKind::Switch | TokenKind::Case | TokenKind::Default | TokenKind::New |
            TokenKind::Delete | TokenKind::Typeof | TokenKind::Void | TokenKind::Instanceof |
            TokenKind::In | TokenKind::Of | TokenKind::This | TokenKind::Super |
            TokenKind::Class | TokenKind::Extends | TokenKind::Static | TokenKind::Get |
            TokenKind::Set | TokenKind::Throw | TokenKind::Try | TokenKind::Catch |
            TokenKind::Finally | TokenKind::True | TokenKind::False | TokenKind::Null |
            TokenKind::Undefined | TokenKind::Import | TokenKind::Export | TokenKind::From |
            TokenKind::As | TokenKind::Async | TokenKind::Await | TokenKind::Yield
        )
    }

    fn eat_semicolon(&mut self) {
        // Automatic semicolon insertion: eat if present, otherwise rely on newline/EOF/}
        self.eat(TokenKind::Semicolon);
    }

    fn prev_followed_by_newline(&self) -> bool {
        if self.pos == 0 || self.pos >= self.tokens.len() {
            return false;
        }
        self.tokens[self.pos - 1].span.line < self.tokens[self.pos].span.line
    }

    fn is_in_call_args(&self) -> bool {
        // Rough heuristic: we're not really tracking this, so just return false.
        // Comma in arguments is handled by parse_arguments directly.
        false
    }

    fn peek_is_method_name(&self) -> bool {
        // Look ahead to see if the next token could be a method name
        if self.pos + 1 < self.tokens.len() {
            matches!(self.tokens[self.pos + 1].kind,
                TokenKind::Identifier(_) | TokenKind::String(_) | TokenKind::Number(_) |
                TokenKind::LBracket
            )
        } else {
            false
        }
    }

    fn could_be_arrow_params(&self) -> bool {
        // Simple check: if we see an identifier, it might be arrow params
        matches!(self.current_kind(), TokenKind::Identifier(_) | TokenKind::DotDotDot)
    }

    fn try_parse_arrow_params(&mut self) -> Result<Option<Vec<Param>>, String> {
        let saved = self.pos;
        let mut params = Vec::new();

        loop {
            let rest = if self.check_kind(&TokenKind::DotDotDot) { self.advance(); true } else { false };
            match self.current_kind() {
                TokenKind::Identifier(name) => {
                    self.advance();
                    let default = if self.check_kind(&TokenKind::Eq) {
                        self.advance();
                        Some(self.parse_assignment_expression()?)
                    } else { None };
                    params.push(Param { name, default, rest });
                }
                _ => {
                    self.pos = saved;
                    return Ok(None);
                }
            }
            if self.check_kind(&TokenKind::RParen) {
                self.advance();
                if self.check_kind(&TokenKind::Arrow) {
                    return Ok(Some(params));
                } else {
                    self.pos = saved;
                    return Ok(None);
                }
            }
            if !self.eat(TokenKind::Comma) {
                self.pos = saved;
                return Ok(None);
            }
        }
    }

    fn expr_to_params(&self, expr: &Expression) -> Option<Vec<Param>> {
        match expr {
            Expression::Identifier(name) => Some(vec![Param::simple(name.clone())]),
            Expression::Sequence(exprs) => {
                let mut params = Vec::new();
                for e in exprs {
                    match e {
                        Expression::Identifier(n) => params.push(Param::simple(n.clone())),
                        _ => return None,
                    }
                }
                Some(params)
            }
            _ => None,
        }
    }

    fn resync_after_template(&mut self) {
        // Re-tokenize from the lexer's current position to get the next token
        // For now, templates with interpolation need special handling in the token stream
        // This is a simplified approach
    }

    fn error(&self, msg: &str) -> String {
        let line = if self.pos < self.tokens.len() {
            self.tokens[self.pos].span.line
        } else if !self.tokens.is_empty() {
            self.tokens.last().unwrap().span.line
        } else {
            1
        };
        format!("Parse error at line {}: {}", line, msg)
    }

    pub fn current_line(&self) -> u32 {
        if self.pos < self.tokens.len() {
            self.tokens[self.pos].span.line
        } else {
            0
        }
    }
}
