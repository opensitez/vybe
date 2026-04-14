//! Load a GrammarDef from a `languages/<lang>/grammar` file.
//!
//! The file format is TOML. See `languages/pascal/grammar` for the canonical example.

use std::collections::HashMap;
use toml::Value;
use crate::grammar::*;

/// Load a grammar from the given file path.
pub fn load_grammar(path: &str) -> Result<GrammarDef, String> {
    let src = std::fs::read_to_string(path)
        .map_err(|e| format!("Cannot read grammar file {}: {}", path, e))?;
    parse_grammar(&src)
}

/// Load a grammar by language name, searching relative to the binary's working directory.
/// Looks for `languages/<lang>/grammar`.
pub fn load_grammar_for(lang: &str) -> Result<GrammarDef, String> {
    // Try relative to cwd, then relative to the binary
    let candidates = [
        format!("languages/{}/grammar", lang),
        format!("../../languages/{}/grammar", lang),
    ];
    for path in &candidates {
        if std::path::Path::new(path).exists() {
            return load_grammar(path);
        }
    }
    Err(format!("Grammar file not found for language '{}'. Looked in languages/{}/grammar", lang, lang))
}

pub fn parse_grammar(src: &str) -> Result<GrammarDef, String> {
    let root: Value = toml::from_str(src)
        .map_err(|e| format!("TOML parse error: {}", e))?;

    let language = parse_language(&root)?;
    let lexer = parse_lexer(&root)?;
    let operators = parse_operator_table(&root)?;
    let blocks = parse_blocks(&root)?;
    let types = parse_types(&root)?;
    let statements = parse_pattern_rules(&root, "statements")?;
    let declarations = parse_pattern_rules(&root, "declarations")?;
    let expressions = parse_expressions(&root)?;
    let params = parse_params(&root)?;
    let assignment = parse_assignment(&root)?;
    let program = parse_program(&root)?;

    Ok(GrammarDef {
        language,
        lexer,
        operators,
        blocks,
        types,
        statements,
        declarations,
        expressions,
        params,
        assignment,
        program,
    })
}

// ── [language] ───────────────────────────────────────────────────────────────

fn parse_language(root: &Value) -> Result<LanguageSpec, String> {
    let t = section(root, "language")?;
    let name = str_field(t, "name")?;
    let case_sensitive = bool_field(t, "case_sensitive").unwrap_or(true);
    let indentation_based = bool_field(t, "indentation_based").unwrap_or(false);
    let expression_language = bool_field(t, "expression_language").unwrap_or(false);

    let statement_terminator = match t.get("statement_terminator")
        .and_then(|v| v.as_str()).unwrap_or(";") {
        ";" => Terminator::Char(';'),
        "newline" | "NEWLINE" => Terminator::Newline,
        "none" | "None" => Terminator::None,
        "asi" | "ASI" => Terminator::Asi,
        other if other.len() == 1 => Terminator::Char(other.chars().next().unwrap()),
        _ => Terminator::Char(';'),
    };

    Ok(LanguageSpec { name, case_sensitive, statement_terminator, indentation_based, expression_language })
}

// ── [lexer] ──────────────────────────────────────────────────────────────────

fn parse_lexer(root: &Value) -> Result<LexerSpec, String> {
    let t = section(root, "lexer")?;

    let comment_line = str_or_str_array(t, "comment_line");
    let comment_block = parse_comment_blocks(t);
    let string_delimiters = str_or_str_array(t, "string_delimiters");
    let string_escape = t.get("string_escape").and_then(|v| v.as_str()).map(|s| s.to_string());
    let triple_string = str_or_str_array(t, "triple_string");
    let string_prefixes = str_or_str_array(t, "string_prefix")
        .into_iter().chain(str_or_str_array(t, "string_prefixes")).collect();
    let interpolation = parse_string_pair(t, "interpolation");
    let template_string = t.get("template_string").and_then(|v| v.as_str()).map(|s| s.to_string());
    let char_prefix = t.get("char_prefix").and_then(|v| v.as_str()).map(|s| s.to_string());
    let hex_prefix = t.get("hex_prefix").and_then(|v| v.as_str()).map(|s| s.to_string());
    let var_prefix = t.get("var_prefix").and_then(|v| v.as_str())
        .and_then(|s| s.chars().next());

    let keywords = str_array(t, "keywords");
    let mut operators = str_array(t, "operators");
    // Sort longest-first for correct lexing
    operators.sort_by(|a, b| b.len().cmp(&a.len()));

    Ok(LexerSpec {
        comment_line,
        comment_block,
        string_delimiters,
        string_escape,
        triple_string,
        string_prefixes,
        interpolation,
        template_string,
        char_prefix,
        hex_prefix,
        var_prefix,
        keywords,
        operators,
    })
}

fn parse_comment_blocks(t: &Value) -> Vec<(String, String)> {
    let mut result = Vec::new();
    if let Some(arr) = t.get("comment_block").and_then(|v| v.as_array()) {
        // Either [["/*", "*/"], ["{", "}"]] or ["/*", "*/"]
        if arr.len() == 2 && arr[0].as_str().is_some() {
            // Single pair: ["/*", "*/"]
            if let (Some(a), Some(b)) = (arr[0].as_str(), arr[1].as_str()) {
                result.push((a.to_string(), b.to_string()));
            }
        } else {
            // Array of pairs
            for item in arr {
                if let Some(pair) = item.as_array() {
                    if pair.len() == 2 {
                        if let (Some(a), Some(b)) = (pair[0].as_str(), pair[1].as_str()) {
                            result.push((a.to_string(), b.to_string()));
                        }
                    }
                }
            }
        }
    }
    result
}

fn parse_string_pair(t: &Value, key: &str) -> Option<(String, String)> {
    t.get(key).and_then(|v| v.as_array()).and_then(|arr| {
        if arr.len() == 2 {
            if let (Some(a), Some(b)) = (arr[0].as_str(), arr[1].as_str()) {
                return Some((a.to_string(), b.to_string()));
            }
        }
        None
    })
}

// ── [operators] ──────────────────────────────────────────────────────────────

fn parse_operator_table(root: &Value) -> Result<OperatorTable, String> {
    let t = match root.get("operators") {
        Some(v) => v,
        None => return Ok(OperatorTable { prefix: vec![], postfix: vec![], infix: vec![] }),
    };

    let prefix = t.get("prefix").and_then(|v| v.as_array())
        .map(|arr| {
            // [[operators.prefix]] has ops = [...]
            // Could be array of tables or direct array of strings
            if let Some(first) = arr.first() {
                if first.as_table().is_some() {
                    // Array of tables: [[operators.prefix]] ops = [...]
                    arr.iter().flat_map(|item| {
                        str_array(item, "ops")
                    }).collect()
                } else {
                    // Direct string array
                    arr.iter().filter_map(|v| v.as_str().map(|s| s.to_string())).collect()
                }
            } else {
                vec![]
            }
        }).unwrap_or_default();

    let postfix = t.get("postfix").and_then(|v| v.as_array())
        .map(|arr| {
            if let Some(first) = arr.first() {
                if first.as_table().is_some() {
                    arr.iter().flat_map(|item| str_array(item, "ops")).collect()
                } else {
                    arr.iter().filter_map(|v| v.as_str().map(|s| s.to_string())).collect()
                }
            } else { vec![] }
        }).unwrap_or_default();

    let infix = t.get("infix").and_then(|v| v.as_array())
        .map(|arr| {
            arr.iter().filter_map(|item| {
                let precedence = item.get("precedence").and_then(|v| v.as_integer()).unwrap_or(1) as u8;
                let ops = str_array(item, "ops");
                let assoc = match item.get("assoc").and_then(|v| v.as_str()).unwrap_or("left") {
                    "right" => Assoc::Right,
                    _ => Assoc::Left,
                };
                Some(InfixLevel { precedence, ops, assoc })
            }).collect()
        }).unwrap_or_default();

    Ok(OperatorTable { prefix, postfix, infix })
}

// ── [blocks] ─────────────────────────────────────────────────────────────────

fn parse_blocks(root: &Value) -> Result<BlockSpec, String> {
    let t = match root.get("blocks") {
        Some(v) => v,
        None => return Ok(BlockSpec { open: "{".into(), close: "}".into(), prefix: None, close_with_kind: false }),
    };

    let (open, close, close_with_kind) = if let Some(compound) = t.get("compound") {
        let open = compound.get("open").and_then(|v| v.as_str()).unwrap_or("{").to_string();
        let close = compound.get("close").and_then(|v| v.as_str()).unwrap_or("}").to_string();
        let cwk = compound.get("close_with_kind").and_then(|v| v.as_bool()).unwrap_or(false);
        (open, close, cwk)
    } else {
        ("{".to_string(), "}".to_string(), false)
    };

    let prefix = t.get("compound").and_then(|c| c.get("prefix"))
        .and_then(|v| v.as_str()).map(|s| s.to_string());

    Ok(BlockSpec { open, close, prefix, close_with_kind })
}

// ── [types] ──────────────────────────────────────────────────────────────────

fn parse_types(root: &Value) -> Result<TypeSpec, String> {
    let t = match root.get("types") {
        Some(v) => v,
        None => return Ok(TypeSpec { position: TypePosition::None, separator: None, return_separator: None }),
    };

    let position = match t.get("position").and_then(|v| v.as_str()).unwrap_or("none") {
        "before" => TypePosition::Before,
        "after" => TypePosition::After,
        _ => TypePosition::None,
    };
    let separator = t.get("separator").and_then(|v| v.as_str()).map(|s| s.to_string());
    let return_separator = t.get("return_separator").and_then(|v| v.as_str()).map(|s| s.to_string());

    Ok(TypeSpec { position, separator, return_separator })
}

// ── [statements] / [declarations] ────────────────────────────────────────────

fn parse_pattern_rules(root: &Value, section_name: &str) -> Result<Vec<PatternRule>, String> {
    let t = match root.get(section_name) {
        Some(v) => v,
        None => return Ok(vec![]),
    };

    let table = match t.as_table() {
        Some(t) => t,
        None => return Ok(vec![]),
    };

    let mut rules = Vec::new();
    for (name, val) in table {
        if let Some(rule_table) = val.as_table() {
            let pattern_str = rule_table.get("pattern")
                .and_then(|v| v.as_str()).unwrap_or("").to_string();
            let maps_to = rule_table.get("maps_to")
                .and_then(|v| v.as_str()).unwrap_or(name).to_string();
            let pattern = parse_pattern_string(&pattern_str);

            let mut extra = HashMap::new();
            if let Some(extra_table) = rule_table.get("extra").and_then(|v| v.as_table()) {
                for (k, v) in extra_table {
                    let val_str = match v {
                        Value::String(s) => s.clone(),
                        Value::Boolean(b) => b.to_string(),
                        Value::Integer(i) => i.to_string(),
                        _ => v.to_string(),
                    };
                    extra.insert(k.clone(), val_str);
                }
            }

            rules.push(PatternRule { name: name.clone(), pattern, maps_to, extra });
        }
    }
    Ok(rules)
}

/// Parse a pattern string like `'"if" EXPR "then" BLOCK ("else" BLOCK)?'`
/// into a Vec<PatternElement>.
fn parse_pattern_string(s: &str) -> Vec<PatternElement> {
    let mut elements = Vec::new();
    let mut chars = s.chars().peekable();

    while let Some(&ch) = chars.peek() {
        match ch {
            ' ' | '\t' | '\n' => { chars.next(); }
            '"' => {
                chars.next();
                let mut kw = String::new();
                while let Some(&c) = chars.peek() {
                    if c == '"' { chars.next(); break; }
                    kw.push(c); chars.next();
                }
                elements.push(PatternElement::Keyword(kw));
            }
            _ => {
                // Read a word
                let mut word = String::new();
                while let Some(&c) = chars.peek() {
                    if c == ' ' || c == '\t' || c == '"' || c == '(' { break; }
                    word.push(c); chars.next();
                }
                match word.as_str() {
                    "EXPR" => elements.push(PatternElement::Expr),
                    "IDENT" => elements.push(PatternElement::Ident),
                    "BLOCK" => elements.push(PatternElement::Block),
                    "STMT_LIST" => elements.push(PatternElement::StmtList),
                    "TYPE" => elements.push(PatternElement::Type),
                    "PARAMS" => elements.push(PatternElement::Params),
                    "IDENT_LIST" => elements.push(PatternElement::IdentList),
                    "EXPR_LIST" => elements.push(PatternElement::ExprList),
                    "DECL_SECTION" => elements.push(PatternElement::DeclSection),
                    "CASE_ARMS" => elements.push(PatternElement::CaseArms),
                    "CATCH_CLAUSES" => elements.push(PatternElement::CatchClauses),
                    "CLASS_MEMBERS" => elements.push(PatternElement::ClassMembers),
                    "INTERFACE_MEMBERS" => elements.push(PatternElement::InterfaceMembers),
                    "RECORD_MEMBERS" => elements.push(PatternElement::RecordMembers),
                    "ENUM_MEMBERS" => elements.push(PatternElement::EnumMembers),
                    "STRING" | "STRING_LIT" => elements.push(PatternElement::StringLit),
                    "NEWLINE" => elements.push(PatternElement::Newline),
                    _ => {} // unknown — skip
                }
            }
        }
    }
    elements
}

// ── [expressions] ────────────────────────────────────────────────────────────

fn parse_expressions(root: &Value) -> Result<ExpressionSpec, String> {
    let t = match root.get("expressions") {
        Some(v) => v,
        None => return Ok(ExpressionSpec {
            member_access: Some(".".into()), optional_chain: None,
            index_open: Some("[".into()), index_close: Some("]".into()),
            call_open: Some("(".into()), call_close: Some(")".into()),
            deref: None, primary_forms: vec![],
        }),
    };

    let member_access = t.get("member_access").and_then(|v| v.as_str()).map(|s| s.to_string());
    let optional_chain = t.get("optional_chain").and_then(|v| v.as_str()).map(|s| s.to_string());
    let deref = t.get("deref").and_then(|v| v.as_str()).map(|s| s.to_string());

    let (index_open, index_close) = parse_string_pair(t, "index_access")
        .map(|(a, b)| (Some(a), Some(b)))
        .unwrap_or((Some("[".into()), Some("]".into())));

    let (call_open, call_close) = parse_string_pair(t, "call")
        .map(|(a, b)| (Some(a), Some(b)))
        .unwrap_or((Some("(".into()), Some(")".into())));

    Ok(ExpressionSpec {
        member_access, optional_chain,
        index_open, index_close,
        call_open, call_close,
        deref, primary_forms: vec![],
    })
}

// ── [params] ─────────────────────────────────────────────────────────────────

fn parse_params(root: &Value) -> Result<ParamSpec, String> {
    let t = match root.get("params") {
        Some(v) => v,
        None => return Ok(ParamSpec {
            open: "(".into(), close: ")".into(), separator: ",".into(),
            name_type_sep: None, type_position: TypePosition::None,
            default_value: None, rest_prefix: None, kwargs_prefix: None,
            multi_name: false, multi_name_sep: None,
            pass_by: HashMap::new(),
        }),
    };

    let open = t.get("open").and_then(|v| v.as_str()).unwrap_or("(").to_string();
    let close = t.get("close").and_then(|v| v.as_str()).unwrap_or(")").to_string();
    let separator = t.get("separator").and_then(|v| v.as_str()).unwrap_or(",").to_string();
    let name_type_sep = t.get("name_type_separator").and_then(|v| v.as_str()).map(|s| s.to_string());
    let type_position = match t.get("type_position").and_then(|v| v.as_str()).unwrap_or("none") {
        "before" => TypePosition::Before,
        "after" => TypePosition::After,
        _ => TypePosition::None,
    };
    let default_value = t.get("default_value").and_then(|v| v.as_str()).map(|s| s.to_string());
    let rest_prefix = t.get("rest_prefix").and_then(|v| v.as_str()).map(|s| s.to_string());
    let kwargs_prefix = t.get("kwargs_prefix").and_then(|v| v.as_str()).map(|s| s.to_string());
    let multi_name = bool_field(t, "multi_name").unwrap_or(false);
    let multi_name_sep = t.get("multi_name_separator").and_then(|v| v.as_str()).map(|s| s.to_string());

    let mut pass_by = HashMap::new();
    if let Some(pb) = t.get("pass_by_keywords").and_then(|v| v.as_table()) {
        for (k, v) in pb {
            if let Some(s) = v.as_str() {
                pass_by.insert(k.clone(), s.to_string());
            }
        }
    }

    Ok(ParamSpec { open, close, separator, name_type_sep, type_position, default_value, rest_prefix, kwargs_prefix, multi_name, multi_name_sep, pass_by })
}

// ── [assignment] ─────────────────────────────────────────────────────────────

fn parse_assignment(root: &Value) -> Result<AssignmentSpec, String> {
    let t = match root.get("assignment") {
        Some(v) => v,
        None => return Ok(AssignmentSpec { operator: Some("=".into()), compound: HashMap::new(), walrus: None }),
    };

    let operator = t.get("operator").and_then(|v| v.as_str()).map(|s| s.to_string());
    let walrus = t.get("walrus").and_then(|v| v.as_str()).map(|s| s.to_string());

    let mut compound = HashMap::new();
    // Support both inline { "+=" = "Add" } and [assignment.compound_operators] section
    let compound_src = t.get("compound_operators").and_then(|v| v.as_table());
    if let Some(ct) = compound_src {
        for (k, v) in ct {
            if let Some(s) = v.as_str() {
                compound.insert(k.clone(), s.to_string());
            }
        }
    }

    Ok(AssignmentSpec { operator, compound, walrus })
}

// ── [program] ────────────────────────────────────────────────────────────────

fn parse_program(root: &Value) -> Result<ProgramSpec, String> {
    // Program section is optional
    Ok(ProgramSpec { header: None, uses: None, body: None })
}

// ── Helpers ──────────────────────────────────────────────────────────────────

fn section<'a>(root: &'a Value, name: &str) -> Result<&'a Value, String> {
    root.get(name).ok_or_else(|| format!("Missing [{}] section", name))
}

fn str_field(t: &Value, key: &str) -> Result<String, String> {
    t.get(key).and_then(|v| v.as_str())
        .map(|s| s.to_string())
        .ok_or_else(|| format!("Missing string field '{}'", key))
}

fn bool_field(t: &Value, key: &str) -> Option<bool> {
    t.get(key).and_then(|v| v.as_bool())
}

fn str_array(t: &Value, key: &str) -> Vec<String> {
    t.get(key)
        .and_then(|v| v.as_array())
        .map(|arr| arr.iter().filter_map(|v| v.as_str().map(|s| s.to_string())).collect())
        .unwrap_or_default()
}

fn str_or_str_array(t: &Value, key: &str) -> Vec<String> {
    match t.get(key) {
        Some(Value::String(s)) => vec![s.clone()],
        Some(Value::Array(_)) => str_array(t, key),
        _ => vec![],
    }
}
