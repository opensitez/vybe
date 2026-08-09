//! Symbol extraction from the common AST.
//!
//! ONE extractor for ALL languages. Each parser converts to
//! `crate::ast::Module`, then this module extracts symbols.

use super::symbols::*;
use crate::ast::*;

/// Extract symbols from a parsed common AST module.
pub fn extract_symbols(module: &Module) -> Vec<Symbol> {
    let mut out = Vec::new();
    for stmt in &module.body {
        extract_stmt(stmt, &mut out);
    }
    out
}

fn format_params(params: &[Param]) -> String {
    params
        .iter()
        .map(|p| {
            if let Some(ref t) = p.type_hint {
                format!("{}: {}", p.name, t)
            } else {
                p.name.clone()
            }
        })
        .collect::<Vec<_>>()
        .join(", ")
}

fn extract_stmt(stmt: &Statement, out: &mut Vec<Symbol>) {
    match &stmt.kind {
        StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            body,
            ..
        } => {
            let param_str = format_params(params);
            let detail = if let Some(rt) = return_type {
                format!("({}): {}", param_str, rt)
            } else {
                format!("({})", param_str)
            };
            let kind = if return_type.is_some() {
                SymbolKind::Function
            } else {
                SymbolKind::Procedure
            };
            let mut children = Vec::new();
            for s in body {
                extract_stmt(s, &mut children);
            }
            out.push(Symbol {
                name: name.clone(),
                kind,
                detail,
                line: stmt.span.start_line,
                end_line: stmt.span.end_line,
                children,
            });
        }

        StmtKind::ClassDecl {
            name,
            parents,
            members,
            ..
        } => {
            let mut children = Vec::new();
            for m in members {
                extract_class_member(m, &mut children);
            }
            let detail = if parents.is_empty() {
                String::new()
            } else {
                format!("({})", parents.join(", "))
            };
            out.push(Symbol {
                name: name.clone(),
                kind: SymbolKind::Class,
                detail,
                line: stmt.span.start_line,
                end_line: stmt.span.end_line,
                children,
            });
        }

        StmtKind::InterfaceDecl { name, members, .. } => {
            let mut children = Vec::new();
            for m in members {
                extract_interface_member(m, &mut children);
            }
            out.push(Symbol {
                name: name.clone(),
                kind: SymbolKind::Interface,
                detail: String::new(),
                line: stmt.span.start_line,
                end_line: stmt.span.end_line,
                children,
            });
        }

        StmtKind::EnumDecl { name, members, .. } => {
            let children: Vec<Symbol> = members
                .iter()
                .map(|m| Symbol {
                    name: m.name.clone(),
                    kind: SymbolKind::EnumMember,
                    detail: String::new(),
                    line: 0,
                    end_line: 0,
                    children: Vec::new(),
                })
                .collect();
            out.push(Symbol {
                name: name.clone(),
                kind: SymbolKind::Enum,
                detail: String::new(),
                line: stmt.span.start_line,
                end_line: stmt.span.end_line,
                children,
            });
        }

        StmtKind::StructDecl { name, members, .. } => {
            let mut children = Vec::new();
            for m in members {
                extract_class_member(m, &mut children);
            }
            out.push(Symbol {
                name: name.clone(),
                kind: SymbolKind::Struct,
                detail: String::new(),
                line: stmt.span.start_line,
                end_line: stmt.span.end_line,
                children,
            });
        }

        StmtKind::ModuleDecl { name, members, .. } => {
            let mut children = Vec::new();
            for m in members {
                extract_class_member(m, &mut children);
            }
            out.push(Symbol {
                name: name.clone(),
                kind: SymbolKind::Module,
                detail: String::new(),
                line: stmt.span.start_line,
                end_line: stmt.span.end_line,
                children,
            });
        }

        StmtKind::NamespaceDecl { name, body } => {
            let mut children = Vec::new();
            for s in body {
                extract_stmt(s, &mut children);
            }
            out.push(Symbol {
                name: name.clone(),
                kind: SymbolKind::Module,
                detail: String::new(),
                line: stmt.span.start_line,
                end_line: stmt.span.end_line,
                children,
            });
        }

        StmtKind::VarDecl { declarations, kind } => {
            let is_const = matches!(kind, VarDeclKind::Const);
            let sym_kind = if is_const {
                SymbolKind::Constant
            } else {
                SymbolKind::Variable
            };
            for decl in declarations {
                let name = match &decl.pattern {
                    BindingPattern::Ident(n) => n.clone(),
                    _ => continue,
                };
                let detail = decl.type_hint.as_deref().unwrap_or_default().to_string();
                out.push(Symbol {
                    name,
                    kind: sym_kind,
                    detail,
                    line: stmt.span.start_line,
                    end_line: stmt.span.end_line,
                    children: Vec::new(),
                });
            }
        }

        StmtKind::DelegateDecl {
            name,
            params,
            return_type,
            ..
        } => {
            let param_str = format_params(params);
            let detail = if let Some(rt) = return_type {
                format!("({}): {}", param_str, rt)
            } else {
                format!("({})", param_str)
            };
            out.push(Symbol {
                name: name.clone(),
                kind: SymbolKind::Type,
                detail,
                line: stmt.span.start_line,
                end_line: stmt.span.end_line,
                children: Vec::new(),
            });
        }

        // Recurse into blocks and control flow to find nested declarations
        StmtKind::Block(stmts) => {
            for s in stmts {
                extract_stmt(s, out);
            }
        }

        // Skip non-declaration statements
        _ => {}
    }
}

fn extract_class_member(member: &ClassMember, out: &mut Vec<Symbol>) {
    match member {
        ClassMember::Field {
            name, type_hint, ..
        } => {
            out.push(Symbol {
                name: name.clone(),
                kind: SymbolKind::Field,
                detail: type_hint.clone().unwrap_or_default(),
                line: 0,
                end_line: 0,
                children: Vec::new(),
            });
        }
        ClassMember::Method(stmt) => {
            extract_stmt(stmt, out);
        }
        ClassMember::Constructor { params, .. } => {
            let param_str = format_params(params);
            out.push(Symbol {
                name: "New".to_string(),
                kind: SymbolKind::Constructor,
                detail: format!("({})", param_str),
                line: 0,
                end_line: 0,
                children: Vec::new(),
            });
        }
        ClassMember::Property {
            name, type_hint, ..
        } => {
            out.push(Symbol {
                name: name.clone(),
                kind: SymbolKind::Property,
                detail: type_hint.clone().unwrap_or_default(),
                line: 0,
                end_line: 0,
                children: Vec::new(),
            });
        }
        ClassMember::Event {
            name, type_hint, ..
        } => {
            out.push(Symbol {
                name: name.clone(),
                kind: SymbolKind::Event,
                detail: type_hint.clone().unwrap_or_default(),
                line: 0,
                end_line: 0,
                children: Vec::new(),
            });
        }
        ClassMember::Const {
            name, type_hint, ..
        } => {
            out.push(Symbol {
                name: name.clone(),
                kind: SymbolKind::Constant,
                detail: type_hint.clone().unwrap_or_default(),
                line: 0,
                end_line: 0,
                children: Vec::new(),
            });
        }
        ClassMember::NestedType(stmt) => {
            extract_stmt(stmt, out);
        }
        // `use T;` / `with M` declares no symbol of its own — the members it
        // contributes belong to the augmenting type, which is outlined where it
        // is declared.
        ClassMember::Augment(_) => {}
    }
}

fn extract_interface_member(member: &InterfaceMember, out: &mut Vec<Symbol>) {
    match member {
        InterfaceMember::Method {
            name,
            params,
            return_type,
            signature_source,
            ..
        } => {
            let param_str = format_params(params);
            let detail = if let Some(rt) = return_type {
                format!("({}): {}", param_str, rt)
            } else if let Some(source) = signature_source {
                format!("({}) -> {}", param_str, source)
            } else {
                format!("({})", param_str)
            };
            let kind = if return_type.is_some() {
                SymbolKind::Function
            } else {
                SymbolKind::Procedure
            };
            out.push(Symbol {
                name: name.clone(),
                kind,
                detail,
                line: 0,
                end_line: 0,
                children: Vec::new(),
            });
        }
        InterfaceMember::Property {
            name, type_hint, ..
        } => {
            out.push(Symbol {
                name: name.clone(),
                kind: SymbolKind::Property,
                detail: type_hint.clone().unwrap_or_default(),
                line: 0,
                end_line: 0,
                children: Vec::new(),
            });
        }
        InterfaceMember::Event {
            name, type_hint, ..
        } => {
            out.push(Symbol {
                name: name.clone(),
                kind: SymbolKind::Event,
                detail: type_hint.clone().unwrap_or_default(),
                line: 0,
                end_line: 0,
                children: Vec::new(),
            });
        }
    }
}

/// Get language keywords for completion.
pub fn language_keywords(lang: Lang) -> &'static [&'static str] {
    match lang {
        Lang::VB => &[
            "Dim",
            "Public",
            "Private",
            "Protected",
            "Friend",
            "Sub",
            "Function",
            "Class",
            "Module",
            "End",
            "If",
            "Then",
            "Else",
            "ElseIf",
            "For",
            "To",
            "Step",
            "Next",
            "Do",
            "While",
            "Loop",
            "Select",
            "Case",
            "Try",
            "Catch",
            "Finally",
            "Throw",
            "Return",
            "Exit",
            "Imports",
            "Namespace",
            "Inherits",
            "Implements",
            "Interface",
            "Enum",
            "Structure",
            "Property",
            "Get",
            "Set",
            "New",
            "Nothing",
            "True",
            "False",
            "And",
            "Or",
            "Not",
            "Is",
            "IsNot",
            "Mod",
            "Like",
            "String",
            "Integer",
            "Double",
            "Boolean",
            "Object",
            "Date",
            "Byte",
            "Long",
            "Short",
            "Console",
            "WriteLn",
            "ReadLine",
            "MsgBox",
            "MessageBox",
        ],
        Lang::JavaScript => &[
            "var",
            "let",
            "const",
            "function",
            "class",
            "if",
            "else",
            "for",
            "while",
            "do",
            "switch",
            "case",
            "break",
            "continue",
            "return",
            "try",
            "catch",
            "finally",
            "throw",
            "new",
            "this",
            "super",
            "extends",
            "import",
            "export",
            "default",
            "async",
            "await",
            "typeof",
            "instanceof",
            "in",
            "of",
            "delete",
            "void",
            "yield",
            "null",
            "undefined",
            "true",
            "false",
            "console",
            "document",
            "window",
            "Array",
            "Object",
            "String",
            "Number",
            "Math",
            "JSON",
            "Promise",
            "Map",
            "Set",
            "Date",
            "RegExp",
            "Error",
        ],
        Lang::CSharp => &[
            "using",
            "namespace",
            "class",
            "struct",
            "interface",
            "enum",
            "public",
            "private",
            "protected",
            "internal",
            "static",
            "void",
            "int",
            "string",
            "bool",
            "double",
            "float",
            "var",
            "if",
            "else",
            "for",
            "foreach",
            "while",
            "do",
            "switch",
            "case",
            "break",
            "continue",
            "return",
            "try",
            "catch",
            "finally",
            "throw",
            "new",
            "this",
            "base",
            "override",
            "virtual",
            "abstract",
            "sealed",
            "readonly",
            "const",
            "null",
            "true",
            "false",
            "async",
            "await",
            "Console",
            "WriteLine",
            "ReadLine",
            "List",
            "Dictionary",
            "Task",
        ],
        Lang::Python => &[
            "def",
            "class",
            "if",
            "elif",
            "else",
            "for",
            "while",
            "break",
            "continue",
            "return",
            "try",
            "except",
            "finally",
            "raise",
            "with",
            "as",
            "import",
            "from",
            "pass",
            "yield",
            "lambda",
            "global",
            "nonlocal",
            "assert",
            "del",
            "True",
            "False",
            "None",
            "and",
            "or",
            "not",
            "in",
            "is",
            "print",
            "len",
            "range",
            "str",
            "int",
            "float",
            "list",
            "dict",
            "set",
            "tuple",
            "self",
            "__init__",
            "__str__",
            "__repr__",
            "super",
            "type",
            "isinstance",
            "enumerate",
            "zip",
            "map",
            "filter",
        ],
        Lang::Ruby => &[
            "def",
            "end",
            "class",
            "module",
            "if",
            "elsif",
            "else",
            "unless",
            "for",
            "while",
            "until",
            "do",
            "begin",
            "rescue",
            "ensure",
            "raise",
            "return",
            "yield",
            "block_given?",
            "self",
            "true",
            "false",
            "nil",
            "and",
            "or",
            "not",
            "require",
            "include",
            "extend",
            "attr_accessor",
            "attr_reader",
            "attr_writer",
            "puts",
            "print",
            "gets",
            "new",
            "initialize",
            "super",
        ],
        Lang::PHP => &[
            "function",
            "class",
            "interface",
            "trait",
            "extends",
            "implements",
            "public",
            "private",
            "protected",
            "static",
            "abstract",
            "final",
            "const",
            "var",
            "if",
            "else",
            "elseif",
            "for",
            "foreach",
            "while",
            "do",
            "switch",
            "case",
            "break",
            "continue",
            "return",
            "try",
            "catch",
            "finally",
            "throw",
            "new",
            "echo",
            "print",
            "null",
            "true",
            "false",
            "array",
            "string",
            "int",
            "float",
            "bool",
            "self",
            "parent",
            "$this",
            "namespace",
            "use",
        ],
        Lang::Dart => &[
            "class",
            "extends",
            "implements",
            "with",
            "mixin",
            "abstract",
            "enum",
            "typedef",
            "import",
            "export",
            "library",
            "part",
            "if",
            "else",
            "for",
            "while",
            "do",
            "switch",
            "case",
            "break",
            "continue",
            "return",
            "try",
            "catch",
            "finally",
            "throw",
            "rethrow",
            "new",
            "this",
            "super",
            "var",
            "final",
            "const",
            "late",
            "void",
            "int",
            "double",
            "String",
            "bool",
            "List",
            "Map",
            "Set",
            "Future",
            "Stream",
            "async",
            "await",
            "yield",
            "true",
            "false",
            "null",
            "print",
            "dynamic",
            "Function",
            "Type",
        ],
        Lang::Pascal => &[
            "program",
            "unit",
            "uses",
            "begin",
            "end",
            "var",
            "const",
            "type",
            "procedure",
            "function",
            "constructor",
            "destructor",
            "class",
            "record",
            "interface",
            "inherited",
            "override",
            "virtual",
            "if",
            "then",
            "else",
            "for",
            "to",
            "downto",
            "do",
            "while",
            "repeat",
            "until",
            "case",
            "of",
            "try",
            "except",
            "finally",
            "raise",
            "exit",
            "break",
            "continue",
            "with",
            "and",
            "or",
            "not",
            "xor",
            "div",
            "mod",
            "in",
            "is",
            "as",
            "nil",
            "true",
            "false",
            "Result",
            "Self",
            "TObject",
            "Create",
            "Destroy",
            "Free",
            "Integer",
            "String",
            "Boolean",
            "Real",
            "Char",
            "WriteLn",
            "ReadLn",
            "Length",
            "Format",
        ],
        Lang::Cobol => &[
            "IDENTIFICATION",
            "DIVISION",
            "ENVIRONMENT",
            "DATA",
            "PROCEDURE",
            "SECTION",
            "WORKING-STORAGE",
            "LINKAGE",
            "PERFORM",
            "MOVE",
            "ADD",
            "SUBTRACT",
            "MULTIPLY",
            "DIVIDE",
            "COMPUTE",
            "IF",
            "ELSE",
            "END-IF",
            "EVALUATE",
            "WHEN",
            "END-EVALUATE",
            "DISPLAY",
            "ACCEPT",
            "STOP",
            "RUN",
            "CALL",
            "USING",
            "PIC",
            "VALUE",
        ],
        _ => &[],
    }
}

/// Detect language from file extension and return Lang enum.
pub fn detect_language(uri: &str) -> Lang {
    let ext = uri.rsplit('.').next().unwrap_or("");
    match ext.to_lowercase().as_str() {
        "vb" | "bas" | "frm" | "cls" => Lang::VB,
        "js" | "mjs" | "cjs" | "ts" => Lang::JavaScript,
        "cs" => Lang::CSharp,
        "py" | "pyw" => Lang::Python,
        "rb" => Lang::Ruby,
        "php" => Lang::PHP,
        "dart" => Lang::Dart,
        "pas" | "pp" | "dpr" | "lpr" => Lang::Pascal,
        "cob" | "cbl" => Lang::Cobol,
        _ => Lang::Unknown,
    }
}
