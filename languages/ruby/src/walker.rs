use super::{RubyParser, Rule};
use vybe_ast::*;
use pest::Parser;
use pest::iterators::Pair;
use std::cell::RefCell;
use std::collections::HashMap;

#[derive(Clone, Default)]
struct RubyMethodInfo {
    arity: i64,
    param_count: i64,
}

thread_local! {
    static RUBY_METHODS: RefCell<HashMap<String, RubyMethodInfo>> = RefCell::new(HashMap::new());
    static RUBY_ALIASES: RefCell<HashMap<String, String>> = RefCell::new(HashMap::new());
    static RUBY_MODULE_MEMBERS: RefCell<HashMap<String, Vec<ClassMember>>> = RefCell::new(HashMap::new());
}

// ════════════════════════════════════════════════════════════════════════════
// Entry point
// ════════════════════════════════════════════════════════════════════════════

pub fn parse(source: &str) -> Result<Module, String> {
    RUBY_METHODS.with(|methods| methods.borrow_mut().clear());
    RUBY_ALIASES.with(|aliases| aliases.borrow_mut().clear());
    RUBY_MODULE_MEMBERS.with(|modules| modules.borrow_mut().clear());
    let source = source.replace("2>/dev/null", "");
    let source = normalize_percent_array_literals(&source);
    let source = normalize_ruby_dynamic_method_defs(&source);
    let source = normalize_ruby_const_reads(&source);
    let pairs =
        RubyParser::parse(Rule::program, source.as_str()).map_err(|e| format!("Parse error: {}", e))?;

    let mut body = Vec::new();
    let mut imports = Vec::new();

    for top in pairs {
        let inner = match top.as_rule() {
            Rule::program => top.into_inner(),
            Rule::EOI => continue,
            _ => {
                walk_stmt_into(top, &mut body, &mut imports)?;
                continue;
            }
        };
        for pair in inner {
            match pair.as_rule() {
                Rule::EOI | Rule::NEWLINE => continue,
                _ => walk_stmt_into(pair, &mut body, &mut imports)?,
            }
        }
    }

    normalize_consecutive_prints(&mut body);

    Ok(Module {
        name: "main".into(),
        language: Lang::Ruby,
        body,
        imports,
    })
}

fn normalize_ruby_dynamic_method_defs(source: &str) -> String {
    let mut out = source.to_string();
    out = normalize_ruby_eval_method_defs(&out, "module_eval", "module");
    out = normalize_ruby_eval_method_defs(&out, "class_eval", "class");
    out = normalize_ruby_exec_method_defs(&out, "module_exec", "module");
    out = normalize_ruby_exec_method_defs(&out, "class_exec", "class");
    out = normalize_ruby_instance_eval_method_defs(&out);
    out = normalize_ruby_instance_exec_singleton_defs(&out);
    normalize_ruby_define_method_blocks(&out)
}

fn normalize_ruby_eval_method_defs(source: &str, eval_name: &str, decl: &str) -> String {
    let needle = format!(".{}('def ", eval_name);
    let mut out = String::new();
    let mut rest = source;
    while let Some(pos) = rest.find(&needle) {
        let Some(recv_start) = ruby_receiver_start(rest, pos) else {
            out.push_str(&rest[..pos + needle.len()]);
            rest = &rest[pos + needle.len()..];
            continue;
        };
        let def_start = recv_start + rest[recv_start..].find("'def ").unwrap_or(0) + 1;
        let Some(close_rel) = rest[def_start..].find("')") else {
            break;
        };
        let receiver = rest[recv_start..pos].trim();
        let def_src = &rest[def_start..def_start + close_rel];
        out.push_str(&rest[..recv_start]);
        out.push_str(decl);
        out.push(' ');
        out.push_str(receiver);
        out.push_str("; ");
        out.push_str(def_src);
        out.push_str("; end");
        rest = &rest[def_start + close_rel + 2..];
    }
    out.push_str(rest);
    out
}

fn normalize_ruby_exec_method_defs(source: &str, exec_name: &str, decl: &str) -> String {
    let needle = format!(".{}(", exec_name);
    let mut out = String::new();
    let mut rest = source;
    while let Some(pos) = rest.find(&needle) {
        let Some(recv_start) = ruby_receiver_start(rest, pos) else {
            out.push_str(&rest[..pos + needle.len()]);
            rest = &rest[pos + needle.len()..];
            continue;
        };
        let after_args = pos + needle.len();
        let Some(args_end_rel) = rest[after_args..].find(")") else {
            break;
        };
        let block_start_search = after_args + args_end_rel + 1;
        let Some(open_rel) = rest[block_start_search..].find('{') else {
            break;
        };
        let block_start = block_start_search + open_rel + 1;
        let Some(close_rel) = rest[block_start..].find('}') else {
            break;
        };
        let block = &rest[block_start..block_start + close_rel];
        let Some(def_src) = ruby_extract_def_from_block(block) else {
            out.push_str(&rest[..block_start + close_rel + 1]);
            rest = &rest[block_start + close_rel + 1..];
            continue;
        };
        let receiver = rest[recv_start..pos].trim();
        out.push_str(&rest[..recv_start]);
        out.push_str(decl);
        out.push(' ');
        out.push_str(receiver);
        out.push_str("; ");
        out.push_str(def_src);
        out.push_str("; end");
        rest = &rest[block_start + close_rel + 1..];
    }
    out.push_str(rest);
    out
}

fn normalize_ruby_instance_eval_method_defs(source: &str) -> String {
    let needle = ".instance_eval('def ";
    let mut out = String::new();
    let mut rest = source;
    while let Some(pos) = rest.find(needle) {
        let Some(recv_start) = ruby_receiver_start(rest, pos) else {
            out.push_str(&rest[..pos + needle.len()]);
            rest = &rest[pos + needle.len()..];
            continue;
        };
        let def_start = recv_start + rest[recv_start..].find("'def ").unwrap_or(0) + 1;
        let Some(close_rel) = rest[def_start..].find("')") else {
            break;
        };
        let receiver = rest[recv_start..pos].trim();
        let def_src = &rest[def_start..def_start + close_rel];
        out.push_str(&rest[..recv_start]);
        out.push_str(&ruby_singleton_def_from_def(receiver, def_src));
        rest = &rest[def_start + close_rel + 2..];
    }
    out.push_str(rest);
    out
}

fn normalize_ruby_instance_exec_singleton_defs(source: &str) -> String {
    let needle = ".instance_exec(";
    let mut out = String::new();
    let mut rest = source;
    while let Some(pos) = rest.find(needle) {
        let Some(recv_start) = ruby_receiver_start(rest, pos) else {
            out.push_str(&rest[..pos + needle.len()]);
            rest = &rest[pos + needle.len()..];
            continue;
        };
        let args_start = pos + needle.len();
        let Some(args_end_rel) = rest[args_start..].find(')') else {
            break;
        };
        let arg_expr = rest[args_start..args_start + args_end_rel].trim();
        let block_search = args_start + args_end_rel + 1;
        let Some(open_rel) = rest[block_search..].find('{') else {
            break;
        };
        let block_start = block_search + open_rel + 1;
        let block_open = block_start - 1;
        let Some(block_close) = ruby_find_matching_brace(rest, block_open) else {
            break;
        };
        let block = rest[block_start..block_close].trim();
        let Some((block_var, method_name, method_body)) = ruby_extract_define_singleton_method(block) else {
            out.push_str(&rest[..block_close + 1]);
            rest = &rest[block_close + 1..];
            continue;
        };
        let body = if method_body.trim() == block_var {
            arg_expr
        } else {
            method_body.trim()
        };
        let receiver = rest[recv_start..pos].trim();
        out.push_str(&rest[..recv_start]);
        let _ = receiver;
        out.push_str("class Object; def ");
        out.push_str(method_name);
        out.push_str("; ");
        out.push_str(body);
        out.push_str("; end; end");
        rest = &rest[block_close + 1..];
    }
    out.push_str(rest);
    out
}

fn normalize_ruby_define_method_blocks(source: &str) -> String {
    let needle = "define_method(:";
    let mut out = String::new();
    let mut rest = source;
    while let Some(pos) = rest.find(needle) {
        let name_start = pos + needle.len();
        let Some(name_end_rel) = rest[name_start..].find(')') else {
            break;
        };
        let name = rest[name_start..name_start + name_end_rel].trim();
        let block_search = name_start + name_end_rel + 1;
        let Some(open_rel) = rest[block_search..].find('{') else {
            out.push_str(&rest[..block_search]);
            rest = &rest[block_search..];
            continue;
        };
        let body_start = block_search + open_rel + 1;
        let Some(close_rel) = rest[body_start..].find('}') else {
            break;
        };
        let body = rest[body_start..body_start + close_rel].trim();
        out.push_str(&rest[..pos]);
        out.push_str("def ");
        out.push_str(name);
        out.push_str("; ");
        out.push_str(body);
        out.push_str("; end");
        rest = &rest[body_start + close_rel + 1..];
    }
    out.push_str(rest);
    out
}

fn ruby_receiver_start(source: &str, dot_pos: usize) -> Option<usize> {
    let bytes = source.as_bytes();
    let mut start = dot_pos;
    while start > 0 {
        let ch = bytes[start - 1] as char;
        if ch.is_ascii_alphanumeric() || matches!(ch, '_' | '@') {
            start -= 1;
        } else {
            break;
        }
    }
    (start < dot_pos).then_some(start)
}

fn ruby_extract_def_from_block(block: &str) -> Option<&str> {
    let def_start = block.find("def ")?;
    let after_def = &block[def_start..];
    let end_rel = after_def.find("; end")?;
    Some(after_def[..end_rel + 5].trim())
}

fn ruby_find_matching_brace(source: &str, open_idx: usize) -> Option<usize> {
    let bytes = source.as_bytes();
    if bytes.get(open_idx).copied()? != b'{' {
        return None;
    }
    let mut depth = 0usize;
    for (idx, byte) in bytes.iter().enumerate().skip(open_idx) {
        match *byte {
            b'{' => depth += 1,
            b'}' => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    return Some(idx);
                }
            }
            _ => {}
        }
    }
    None
}

fn ruby_singleton_def_from_def(receiver: &str, def_src: &str) -> String {
    let trimmed = def_src.trim();
    if let Some(rest) = trimmed.strip_prefix("def ") {
        let _ = receiver;
        format!(
            "class Object; def {}; end; end",
            rest.trim_end_matches("; end").trim()
        )
    } else {
        trimmed.to_string()
    }
}

fn ruby_extract_define_singleton_method(block: &str) -> Option<(&str, &str, &str)> {
    let block = block.trim();
    let after_pipe = block.strip_prefix('|')?;
    let pipe_end = after_pipe.find('|')?;
    let block_var = after_pipe[..pipe_end].trim();
    let body = after_pipe[pipe_end + 1..].trim();
    let needle = "define_singleton_method(:";
    let method_start = body.find(needle)? + needle.len();
    let method_end_rel = body[method_start..].find(')')?;
    let method_name = body[method_start..method_start + method_end_rel].trim();
    let block_search = method_start + method_end_rel + 1;
    let open_rel = body[block_search..].find('{')?;
    let inner_start = block_search + open_rel + 1;
    let close_rel = body[inner_start..].find('}')?;
    Some((block_var, method_name, body[inner_start..inner_start + close_rel].trim()))
}

fn normalize_percent_array_literals(source: &str) -> String {
    let mut out = String::new();
    let chars: Vec<char> = source.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        if chars[i] == '%' && i + 2 < chars.len() && matches!(chars[i + 1], 'w' | 'W' | 'i' | 'I') {
            let kind = chars[i + 1];
            let open = chars[i + 2];
            let close = match open {
                '(' => ')',
                '[' => ']',
                '{' => '}',
                '<' => '>',
                '/' => '/',
                '|' => '|',
                '!' => '!',
                _ => {
                    out.push(chars[i]);
                    i += 1;
                    continue;
                }
            };
            let mut j = i + 3;
            let mut body = String::new();
            let mut escaped = false;
            while j < chars.len() {
                let ch = chars[j];
                if escaped {
                    body.push('\\');
                    body.push(ch);
                    escaped = false;
                    j += 1;
                    continue;
                }
                if ch == '\\' {
                    escaped = true;
                    j += 1;
                    continue;
                }
                if ch == close {
                    break;
                }
                body.push(ch);
                j += 1;
            }
            if j < chars.len() && chars[j] == close {
                let interpolate = matches!(kind, 'W' | 'I');
                let symbolish = matches!(kind, 'i' | 'I');
                let words = ruby_percent_words(&body, interpolate);
                out.push('[');
                for (idx, word) in words.iter().enumerate() {
                    if idx > 0 {
                        out.push_str(", ");
                    }
                    if interpolate && word.starts_with("#{") && word.ends_with('}') {
                        out.push_str(&word[2..word.len() - 1]);
                    } else if symbolish && is_simple_ruby_symbol_word(word) {
                        out.push(':');
                        out.push_str(word);
                    } else if !interpolate && word.starts_with("#{") && word.ends_with('}') {
                        out.push_str(&ruby_single_quoted(&format!("\\{}", word)));
                    } else if !interpolate {
                        out.push_str(&ruby_single_quoted(word));
                    } else {
                        out.push_str(&ruby_double_quoted(word));
                    }
                }
                out.push(']');
                i = j + 1;
                continue;
            }
        }
        out.push(chars[i]);
        i += 1;
    }
    out
}

fn normalize_ruby_const_reads(source: &str) -> String {
    let chars: Vec<char> = source.chars().collect();
    let mut out = String::new();
    let mut i = 0;
    let mut in_single = false;
    let mut in_double = false;
    let mut escaped = false;

    while i < chars.len() {
        let ch = chars[i];
        if escaped {
            out.push(ch);
            escaped = false;
            i += 1;
            continue;
        }
        if ch == '\\' && (in_single || in_double) {
            out.push(ch);
            escaped = true;
            i += 1;
            continue;
        }
        if ch == '\'' && !in_double {
            in_single = !in_single;
            out.push(ch);
            i += 1;
            continue;
        }
        if ch == '"' && !in_single {
            in_double = !in_double;
            out.push(ch);
            i += 1;
            continue;
        }
        if ch == '#' && !in_single && !in_double {
            while i < chars.len() {
                out.push(chars[i]);
                if chars[i] == '\n' {
                    break;
                }
                i += 1;
            }
            i += 1;
            continue;
        }
        if !in_single
            && !in_double
            && ch.is_ascii_uppercase()
            && (i == 0 || !(chars[i - 1].is_ascii_alphanumeric() || chars[i - 1] == '_'))
        {
            let start = i;
            let mut j = i + 1;
            while j < chars.len() && (chars[j].is_ascii_alphanumeric() || chars[j] == '_') {
                j += 1;
            }
            if j + 2 < chars.len()
                && chars[j] == ':'
                && chars[j + 1] == ':'
                && chars[j + 2].is_ascii_uppercase()
            {
                let mut k = j + 3;
                while k < chars.len() && (chars[k].is_ascii_alphanumeric() || chars[k] == '_') {
                    k += 1;
                }
                let left: String = chars[start..j].iter().collect();
                let right: String = chars[j + 2..k].iter().collect();
                let prior = out.split_whitespace().last().unwrap_or("");
                let declaration_context = matches!(
                    prior,
                    "class" | "module" | "include" | "extend" | "rescue" | "<"
                );
                if !declaration_context && !(left == "Math" && matches!(right.as_str(), "PI" | "E")) {
                    out.push_str(&left);
                    out.push_str(".const_get(:");
                    out.push_str(&right);
                    out.push(')');
                    i = k;
                    continue;
                }
            }
        }
        out.push(ch);
        i += 1;
    }

    out
}

fn ruby_double_quoted(s: &str) -> String {
    let mut out = String::from("\"");
    for ch in s.chars() {
        match ch {
            '\\' => out.push_str("\\\\"),
            '"' => out.push_str("\\\""),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            _ => out.push(ch),
        }
    }
    out.push('"');
    out
}

fn ruby_single_quoted(s: &str) -> String {
    let mut out = String::from("'");
    for ch in s.chars() {
        match ch {
            '\\' => out.push_str("\\\\"),
            '\'' => out.push_str("\\'"),
            _ => out.push(ch),
        }
    }
    out.push('\'');
    out
}

fn is_simple_ruby_symbol_word(s: &str) -> bool {
    let mut chars = s.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    (first.is_ascii_alphabetic() || first == '_')
        && chars.all(|ch| ch.is_ascii_alphanumeric() || ch == '_')
}

fn walk_stmt_into(
    pair: Pair<Rule>,
    body: &mut Vec<Statement>,
    imports: &mut Vec<Import>,
) -> Result<(), String> {
    match pair.as_rule() {
        Rule::require_stmt => imports.push(walk_require(pair)?),
        _ => {
            let stmt = walk_statement(pair)?;
            if !matches!(stmt.kind, StmtKind::Empty) {
                body.push(stmt);
            }
        }
    }
    Ok(())
}

// ════════════════════════════════════════════════════════════════════════════
// Statements
// ════════════════════════════════════════════════════════════════════════════

fn walk_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::method_def => walk_method_def(pair)?,
        Rule::class_def => walk_class_def(pair)?,
        Rule::module_def => walk_module_def(pair)?,

        Rule::if_stmt => walk_if(pair)?,
        Rule::unless_stmt => walk_unless(pair)?,
        Rule::while_stmt => walk_while(pair)?,
        Rule::until_stmt => walk_until(pair)?,
        Rule::for_stmt => walk_for(pair)?,
        Rule::case_stmt => walk_case(pair)?,
        Rule::begin_stmt => walk_begin(pair)?,
        Rule::loop_stmt => walk_loop(pair)?,

        Rule::return_stmt => walk_return(pair)?,
        Rule::break_stmt => walk_break_or_next(pair, true)?,
        Rule::next_stmt => walk_break_or_next(pair, false)?,
        Rule::raise_stmt => walk_raise(pair)?,
        Rule::retry_stmt => StmtKind::Continue(ContinueTarget::Implicit),
        Rule::redo_stmt => StmtKind::Continue(ContinueTarget::Implicit),

        Rule::require_stmt => return Ok(Statement::new(StmtKind::Empty)), // handled in walk_stmt_into
        Rule::at_exit_stmt => StmtKind::Empty,                            // no runtime equivalent
        Rule::catch_throw_stmt => StmtKind::Empty,                        // simplified
        Rule::access_modifier_stmt => StmtKind::Empty,                    // metadata only
        Rule::alias_stmt => walk_alias_stmt(pair)?,
        Rule::undef_stmt => StmtKind::Empty, // not directly representable

        Rule::multi_assign_stmt => walk_multi_assign(pair)?,
        Rule::expr_or_assign_stmt => walk_expr_or_assign(pair)?,

        Rule::NEWLINE => StmtKind::Empty,

        other => return Err(format!("Unexpected statement rule: {:?}", other)),
    };
    Ok(Statement::with_span(kind, span))
}

// ── Method def ──────────────────────────────────────────────────────────────

fn walk_method_def(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut is_self_method = false;
    let mut params = Vec::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::method_name => {
                let text = p.as_str();
                if text.starts_with("self.") {
                    is_self_method = true;
                    name = text[5..].to_string();
                } else if let Some((_, method)) = text.rsplit_once('.') {
                    name = method.to_string();
                } else {
                    name = text.to_string();
                }
            }
            Rule::method_params => params = walk_method_params(p)?,
            Rule::body => body = walk_body(p)?,
            _ => {}
        }
    }

    // Don't apply implicit return to constructors — the compiler handles constructor return
    if name != "initialize" {
        apply_implicit_return(&mut body);
    }

    let mut modifiers = Modifiers::default();
    if is_self_method {
        modifiers.is_static = true;
    }

    let is_generator = body_has_yield(&body);
    register_ruby_method("Object", &name, &params);

    Ok(StmtKind::FunctionDecl {
        name,
        params,
        return_type: None,
        body,
        modifiers,
        handles: Vec::new(),
        is_async: false,
        is_generator,
        is_sub: false,
    })
}

fn walk_alias_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let names = pair
        .into_inner()
        .filter(|p| p.as_rule() == Rule::method_name_id)
        .map(|p| p.as_str().to_string())
        .collect::<Vec<_>>();
    if names.len() >= 2 {
        RUBY_ALIASES.with(|aliases| {
            aliases
                .borrow_mut()
                .insert(names[0].clone(), names[1].clone());
        });
    }
    Ok(StmtKind::Empty)
}

fn method_key(owner: &str, name: &str) -> String {
    format!("{}::{}", owner, name)
}

fn register_ruby_method(owner: &str, name: &str, params: &[Param]) {
    let arity = params
        .iter()
        .filter(|p| !p.is_optional && !p.is_rest && !p.is_kwargs)
        .count() as i64;
    let param_count = params.len() as i64;
    RUBY_METHODS.with(|methods| {
        methods.borrow_mut().insert(
            method_key(owner, name),
            RubyMethodInfo { arity, param_count },
        );
    });
}

fn register_ruby_module_members(name: &str, members: &[ClassMember]) {
    RUBY_MODULE_MEMBERS.with(|modules| {
        modules
            .borrow_mut()
            .insert(name.to_string(), members.to_vec());
    });
}

fn ruby_module_members(name: &str) -> Vec<ClassMember> {
    RUBY_MODULE_MEMBERS.with(|modules| {
        modules
            .borrow()
            .get(name)
            .cloned()
            .unwrap_or_default()
    })
}

fn register_ruby_member_methods(owner: &str, members: &[ClassMember]) {
    for member in members {
        if let ClassMember::Method(method) = member {
            if let StmtKind::FunctionDecl { name, params, .. } = &method.kind {
                register_ruby_method(owner, name, params);
            }
        }
    }
}

fn ruby_alias_original(name: &str) -> String {
    RUBY_ALIASES.with(|aliases| {
        aliases
            .borrow()
            .get(name)
            .cloned()
            .unwrap_or_else(|| name.to_string())
    })
}

fn ruby_method_info(owner: &str, name: &str) -> RubyMethodInfo {
    let original = ruby_alias_original(name);
    RUBY_METHODS.with(|methods| {
        let methods = methods.borrow();
        methods
            .get(&method_key(owner, name))
            .or_else(|| methods.get(&method_key(owner, &original)))
            .or_else(|| methods.get(&method_key("Object", name)))
            .or_else(|| methods.get(&method_key("Object", &original)))
            .cloned()
            .unwrap_or_default()
    })
}

fn walk_method_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::param_list {
            params = walk_param_list(p)?;
        }
    }
    Ok(params)
}

fn walk_param_list(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::param_item {
            let inner = p.into_inner().next();
            if let Some(item) = inner {
                match item.as_rule() {
                    Rule::normal_param => {
                        params.push(Param {
                            name: item.as_str().to_string(),
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        });
                    }
                    Rule::optional_param => {
                        let mut name = String::new();
                        let mut default = None;
                        for c in item.into_inner() {
                            match c.as_rule() {
                                Rule::identifier => name = c.as_str().to_string(),
                                _ => default = Some(walk_expression(c)?),
                            }
                        }
                        params.push(Param {
                            name,
                            type_hint: None,
                            default,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: true,
                            is_nullable: false,
                        });
                    }
                    Rule::splat_param => {
                        let name = item
                            .into_inner()
                            .find(|c| c.as_rule() == Rule::identifier)
                            .map(|c| c.as_str().to_string())
                            .unwrap_or_default();
                        params.push(Param {
                            name,
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: true,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        });
                    }
                    Rule::double_splat_param => {
                        let name = item
                            .into_inner()
                            .find(|c| c.as_rule() == Rule::identifier)
                            .map(|c| c.as_str().to_string())
                            .unwrap_or_default();
                        params.push(Param {
                            name,
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: true,
                            is_optional: false,
                            is_nullable: false,
                        });
                    }
                    Rule::block_param => {
                        // &block — ignore for now, blocks are handled differently
                    }
                    Rule::keyword_param => {
                        let mut name = String::new();
                        let mut default = None;
                        for c in item.into_inner() {
                            match c.as_rule() {
                                Rule::identifier => name = c.as_str().to_string(),
                                _ if is_expression_rule(c.as_rule()) => {
                                    default = Some(walk_expression(c)?);
                                }
                                _ => {}
                            }
                        }
                        let is_optional = default.is_some();
                        params.push(Param {
                            name,
                            type_hint: None,
                            default,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional,
                            is_nullable: false,
                        });
                    }
                    _ => {}
                }
            }
        }
    }
    Ok(params)
}

// ── Class def ───────────────────────────────────────────────────────────────

fn walk_class_def(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut parents = Vec::new();
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::constant => {
                if name.is_empty() {
                    name = p.as_str().to_string();
                }
            }
            Rule::constant_path => {
                parents.push(p.as_str().to_string());
            }
            Rule::class_body => {
                members = walk_class_body(p, &name)?;
            }
            _ => {}
        }
    }

    Ok(StmtKind::ClassDecl {
        name,
        parents,
        interfaces: Vec::new(),
        members,
        modifiers: ClassModifiers::default(),
        decorators: vec![],
    })
}

fn walk_class_body(pair: Pair<Rule>, class_name: &str) -> Result<Vec<ClassMember>, String> {
    let mut members = Vec::new();
    let mut current_visibility = Visibility::Public;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::access_modifier_stmt => {
                let text = p.as_str().trim();
                if text.starts_with("private") {
                    current_visibility = Visibility::Private;
                } else if text.starts_with("protected") {
                    current_visibility = Visibility::Protected;
                } else {
                    current_visibility = Visibility::Public;
                }
            }
            Rule::attr_decl => {
                members.extend(walk_attr_decl(p)?);
            }
            Rule::method_def => {
                let stmt_kind = walk_method_def(p)?;
                if let StmtKind::FunctionDecl {
                    name,
                    params,
                    body,
                    modifiers,
                    ..
                } = &stmt_kind
                {
                    register_ruby_method(class_name, name, params);
                    if name == "initialize" {
                        // Extract instance variable assignments from constructor body
                        members.push(ClassMember::Constructor {
                            // Ruby has one constructor, `initialize` — unnamed.
                            name: None,
                            params: params.clone(),
                            body: body.clone(),
                            base_args: None,
                            initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
                            visibility: current_visibility,
                        });
                    } else {
                        let mut mods = modifiers.clone();
                        mods.visibility = current_visibility;
                        members.push(ClassMember::Method(Box::new(Statement::new(stmt_kind))));
                    }
                }
            }
            Rule::include_stmt | Rule::extend_stmt => {
                let included = p
                    .into_inner()
                    .find(|inner| matches!(inner.as_rule(), Rule::constant_path | Rule::constant))
                    .map(|inner| inner.as_str().to_string());
                if let Some(module_name) = included {
                    let module_members = ruby_module_members(&module_name);
                    register_ruby_member_methods(class_name, &module_members);
                    members.extend(module_members);
                }
            }
            Rule::alias_stmt => {}
            Rule::class_def => {
                // Nested class
                let nested = walk_class_def(p)?;
                members.push(ClassMember::NestedType(Box::new(Statement::new(nested))));
            }
            Rule::module_def => {
                let nested = walk_module_def(p)?;
                members.push(ClassMember::NestedType(Box::new(Statement::new(nested))));
            }
            Rule::NEWLINE => {}
            _ => {
                // Other statements in class body → treat as static initializer
                let stmt = walk_statement(p)?;
                if !matches!(stmt.kind, StmtKind::Empty) {
                    if let Some((alias, original)) = ruby_alias_method_stmt(&stmt) {
                        if let Some(alias_stmt) = members.iter().find_map(|member| {
                            let ClassMember::Method(method) = member else {
                                return None;
                            };
                            let mut cloned = (**method).clone();
                            if let StmtKind::FunctionDecl { name, .. } = &mut cloned.kind {
                                if name == &original {
                                    *name = alias.clone();
                                    return Some(cloned);
                                }
                            }
                            None
                        }) {
                            members.push(ClassMember::Method(Box::new(alias_stmt)));
                        }
                        continue;
                    }
                    if let Some(name) = ruby_remove_const_stmt(&stmt) {
                        members.retain(|member| {
                            !matches!(member, ClassMember::Const { name: const_name, .. } if const_name == &name)
                        });
                        continue;
                    }
                    if let Some(name) = ruby_remove_method_stmt(&stmt) {
                        members.retain(|member| {
                            !matches!(
                                member,
                                ClassMember::Method(method)
                                    if matches!(&method.kind, StmtKind::FunctionDecl { name: method_name, .. } if method_name == &name)
                            )
                        });
                        continue;
                    }
                    if let StmtKind::Assign { targets, value } = stmt.kind {
                        if targets.len() == 1 {
                            if let ExprKind::Ident(name) = &targets[0].kind {
                                if name.chars().next().is_some_and(|c| c.is_ascii_uppercase()) {
                                    members.push(ClassMember::Const {
                                        name: name.clone(),
                                        type_hint: None,
                                        value,
                                        visibility: current_visibility,
                                    });
                                    continue;
                                }
                            }
                        }
                        members.push(ClassMember::Method(Box::new(Statement::new(
                            StmtKind::Assign { targets, value },
                        ))));
                    } else {
                        members.push(ClassMember::Method(Box::new(stmt)));
                    }
                }
            }
        }
    }
    Ok(members)
}

fn ruby_remove_const_stmt(stmt: &Statement) -> Option<String> {
    let StmtKind::Expr(expr) = &stmt.kind else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(&callee.kind, ExprKind::Ident(name) if name == "remove_const") || args.len() != 1 {
        return None;
    }
    ruby_method_name_arg(&args[0].value)
}

fn ruby_remove_method_stmt(stmt: &Statement) -> Option<String> {
    let StmtKind::Expr(expr) = &stmt.kind else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(&callee.kind, ExprKind::Ident(name) if name == "remove_method") || args.len() != 1 {
        return None;
    }
    ruby_method_name_arg(&args[0].value)
}

fn ruby_alias_method_stmt(stmt: &Statement) -> Option<(String, String)> {
    let StmtKind::Expr(expr) = &stmt.kind else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(&callee.kind, ExprKind::Ident(name) if name == "alias_method") || args.len() != 2 {
        return None;
    }
    Some((
        ruby_method_name_arg(&args[0].value)?,
        ruby_method_name_arg(&args[1].value)?,
    ))
}

fn walk_attr_decl(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let mut kind = "";
    let mut names = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::attr_kind => {
                kind = match p.as_str().trim() {
                    "attr_accessor" => "accessor",
                    "attr_reader" => "reader",
                    "attr_writer" => "writer",
                    _ => "accessor",
                }
            }
            Rule::symbol_list => {
                for s in p.into_inner() {
                    let text = s.as_str().trim();
                    let name = if text.starts_with(':') {
                        &text[1..]
                    } else if text.starts_with('"') || text.starts_with('\'') {
                        &text[1..text.len() - 1]
                    } else {
                        text
                    };
                    names.push(name.to_string());
                }
            }
            _ => {}
        }
    }

    let mut members = Vec::new();
    for name in names {
        let has_getter = kind == "accessor" || kind == "reader";
        let has_setter = kind == "accessor" || kind == "writer";

        // Getter → method `name()` that returns self._rb_<field>
        // The backing field is created by `@name = ...` in initialize
        // which maps to self._rb_name (prefixed to avoid struct key collision)
        if has_getter {
            let self_expr = Expression::new(ExprKind::Ident("self".into()));
            let field_access = Expression::new(ExprKind::Member {
                object: Box::new(self_expr),
                field: format!("_rb_{}", name),
                null_safe: false,
            });
            let body = vec![Statement::new(StmtKind::Return(Some(field_access)))];
            members.push(ClassMember::Method(Box::new(Statement::new(
                StmtKind::FunctionDecl {
                    name: name.clone(),
                    params: Vec::new(),
                    return_type: None,
                    body,
                    modifiers: Modifiers::default(),
                    handles: Vec::new(),
                    is_async: false,
                    is_generator: false,
                    is_sub: false,
                },
            ))));
        } else if kind == "writer" {
            let body = vec![Statement::new(StmtKind::Throw {
                expr: Some(Expression::string("NoMethodError")),
                cause: None,
            })];
            members.push(ClassMember::Method(Box::new(Statement::new(
                StmtKind::FunctionDecl {
                    name: name.clone(),
                    params: Vec::new(),
                    return_type: None,
                    body,
                    modifiers: Modifiers::default(),
                    handles: Vec::new(),
                    is_async: false,
                    is_generator: false,
                    is_sub: false,
                },
            ))));
        }

        // Setter semantics: Ruby `d.name = x` is transformed in the walker to
        // Assign(Member(d, "_rb_name"), x) via fixup_assign_target, which writes
        // directly to the _rb_ prefixed backing field via struct_set.
        let _ = has_setter;
    }
    Ok(members)
}

// ── Module def ──────────────────────────────────────────────────────────────

fn walk_module_def(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::constant => {
                if name.is_empty() {
                    name = p.as_str().to_string();
                }
            }
            Rule::class_body => {
                members = walk_class_body(p, &name)?;
            }
            _ => {}
        }
    }

    register_ruby_module_members(&name, &members);

    Ok(StmtKind::ModuleDecl {
        name,
        members,
        visibility: Visibility::Public,
    })
}

// ── If ──────────────────────────────────────────────────────────────────────

fn walk_if(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Check for modifier form: expression if_kw expression
    if children.iter().any(|p| p.as_rule() == Rule::if_kw) {
        let mut iter = children.into_iter();
        let body_expr = walk_expression(iter.next().ok_or("Missing modifier if body")?)?;
        // skip if_kw
        iter.find(|p| p.as_rule() == Rule::if_kw);
        let cond = walk_expression(iter.next().ok_or("Missing modifier if condition")?)?;
        return Ok(StmtKind::If {
            cond,
            then_body: vec![Statement::new(StmtKind::Expr(body_expr))],
            elifs: Vec::new(),
            else_body: None,
        });
    }

    // Block form: if cond then_kw? body elsif* else? end
    let mut iter = children.into_iter();
    let cond = walk_expression(next_meaningful(&mut iter)?)?;
    let then_body = walk_body(next_rule(&mut iter, Rule::body)?)?;

    let mut elifs = Vec::new();
    let mut else_body = None;

    for p in iter {
        match p.as_rule() {
            Rule::elsif_clause => {
                let mut ei = p.into_inner();
                let econd = walk_expression(next_meaningful(&mut ei)?)?;
                let ebody = walk_body(find_rule(ei, Rule::body)?)?;
                elifs.push((econd, ebody));
            }
            Rule::else_clause => {
                let ei = p.into_inner();
                else_body = Some(walk_body(find_rule(ei, Rule::body)?)?);
            }
            _ => {}
        }
    }

    Ok(StmtKind::If {
        cond,
        then_body,
        elifs,
        else_body,
    })
}

// ── Unless ──────────────────────────────────────────────────────────────────

fn walk_unless(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Check for modifier form: expression unless_kw expression
    if children.iter().any(|p| p.as_rule() == Rule::unless_kw) {
        let mut iter = children.into_iter();
        let body_expr = walk_expression(iter.next().ok_or("Missing modifier unless body")?)?;
        iter.find(|p| p.as_rule() == Rule::unless_kw);
        let cond = walk_expression(iter.next().ok_or("Missing modifier unless condition")?)?;
        // unless → if !cond
        return Ok(StmtKind::If {
            cond: negate(cond),
            then_body: vec![Statement::new(StmtKind::Expr(body_expr))],
            elifs: Vec::new(),
            else_body: None,
        });
    }

    // Block form
    let mut iter = children.into_iter();
    let cond = walk_expression(next_meaningful(&mut iter)?)?;
    let then_body = walk_body(find_rule_from_iter(&mut iter, Rule::body)?)?;

    let mut else_body = None;
    for p in iter {
        if p.as_rule() == Rule::else_clause {
            let ei = p.into_inner();
            else_body = Some(walk_body(find_rule(ei, Rule::body)?)?);
        }
    }

    // unless cond → if !cond
    Ok(StmtKind::If {
        cond: negate(cond),
        then_body,
        elifs: Vec::new(),
        else_body,
    })
}

// ── While ───────────────────────────────────────────────────────────────────

fn walk_while(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Modifier form: expression while_kw expression
    if children.iter().any(|p| p.as_rule() == Rule::while_kw) {
        let mut iter = children.into_iter();
        let body_expr = walk_expression(iter.next().ok_or("Missing modifier while body")?)?;
        iter.find(|p| p.as_rule() == Rule::while_kw);
        let cond = walk_expression(iter.next().ok_or("Missing modifier while condition")?)?;
        return Ok(StmtKind::While {
            cond,
            body: vec![Statement::new(StmtKind::Expr(body_expr))],
            else_body: None,
        });
    }

    // Block form
    let mut iter = children.into_iter();
    let cond = walk_expression(next_meaningful(&mut iter)?)?;
    let body = walk_body(find_rule_from_iter(&mut iter, Rule::body)?)?;
    Ok(StmtKind::While {
        cond,
        body,
        else_body: None,
    })
}

// ── Until ───────────────────────────────────────────────────────────────────

fn walk_until(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Modifier form
    if children.iter().any(|p| p.as_rule() == Rule::until_kw) {
        let mut iter = children.into_iter();
        let body_expr = walk_expression(iter.next().ok_or("Missing modifier until body")?)?;
        iter.find(|p| p.as_rule() == Rule::until_kw);
        let cond = walk_expression(iter.next().ok_or("Missing modifier until condition")?)?;
        // until cond → while !cond
        return Ok(StmtKind::While {
            cond: negate(cond),
            body: vec![Statement::new(StmtKind::Expr(body_expr))],
            else_body: None,
        });
    }

    // Block form: until cond → while !cond
    let mut iter = children.into_iter();
    let cond = walk_expression(next_meaningful(&mut iter)?)?;
    let body = walk_body(find_rule_from_iter(&mut iter, Rule::body)?)?;
    Ok(StmtKind::While {
        cond: negate(cond),
        body,
        else_body: None,
    })
}

// ── For ─────────────────────────────────────────────────────────────────────

fn walk_for(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut vars = Vec::new();
    let mut iter_expr = None;
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => vars.push(p.as_str().to_string()),
            Rule::in_kw | Rule::do_kw => {}
            Rule::body => body = walk_body(p)?,
            _ if is_expression_rule(p.as_rule()) => {
                if iter_expr.is_none() {
                    iter_expr = Some(walk_expression(p)?);
                }
            }
            _ => {}
        }
    }

    // Multi-target destructuring
    let var = if vars.len() > 1 {
        let tmp = "__forin_element".to_string();
        let mut destructure_stmts: Vec<Statement> = Vec::new();
        for (i, name) in vars.iter().enumerate() {
            destructure_stmts.push(Statement::new(StmtKind::Assign {
                targets: vec![Expression::new(ExprKind::Ident(name.clone()))],
                value: Expression::new(ExprKind::Index {
                    object: Box::new(Expression::new(ExprKind::Ident(tmp.clone()))),
                    index: Box::new(Expression::int(i as i64)),
                    null_safe: false,
                }),
            }));
        }
        destructure_stmts.extend(body);
        body = destructure_stmts;
        tmp
    } else {
        vars.into_iter().next().unwrap_or_default()
    };

    Ok(StmtKind::ForIn {
        var,
        key: None,
        iter: iter_expr.unwrap_or(Expression::null()),
        body,
        of: true,
        else_body: None,
        is_async: false,
    })
}

// ── Case / When ─────────────────────────────────────────────────────────────

fn walk_case(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut subject = None;
    let mut cases = Vec::new();
    let mut default = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::when_clause => {
                let mut conditions = Vec::new();
                let mut body = Vec::new();
                for wp in p.into_inner() {
                    match wp.as_rule() {
                        Rule::expression_list => {
                            for ep in wp.into_inner() {
                                if is_expression_rule(ep.as_rule()) {
                                    let expr = walk_expression(ep)?;
                                    conditions.push(CaseCondition::Value(expr));
                                }
                            }
                        }
                        Rule::body => body = walk_body(wp)?,
                        Rule::then_kw => {}
                        _ if is_expression_rule(wp.as_rule()) => {
                            let expr = walk_expression(wp)?;
                            conditions.push(CaseCondition::Value(expr));
                        }
                        _ => {}
                    }
                }
                cases.push(SwitchCase { conditions, body });
            }
            Rule::else_clause => {
                let ei = p.into_inner();
                default = Some(walk_body(find_rule(ei, Rule::body)?)?);
            }
            _ if is_expression_rule(p.as_rule()) => {
                if subject.is_none() {
                    subject = Some(walk_expression(p)?);
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::Switch {
        expr: subject.unwrap_or(Expression::bool(true)),
        cases,
        default,
    })
}

// ── Begin / Rescue / Ensure ─────────────────────────────────────────────────

fn walk_begin(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body = Vec::new();
    let mut catches = Vec::new();
    let mut else_body = None;
    let mut finally = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::body => {
                if body.is_empty() {
                    body = walk_body(p)?;
                }
            }
            Rule::rescue_clause => {
                let mut types = Vec::new();
                let mut var_name = None;
                let mut catch_body = Vec::new();

                for cp in p.into_inner() {
                    match cp.as_rule() {
                        Rule::constant | Rule::constant_path => types.push(cp.as_str().to_string()),
                        Rule::identifier => var_name = Some(cp.as_str().to_string()),
                        Rule::body => catch_body = walk_body(cp)?,
                        _ => {}
                    }
                }
                catches.push(CatchClause {
                    types,
                    var_name,
                    stack_var: None,
                    body: catch_body,
                    when_clause: None,
                });
            }
            Rule::else_clause => {
                let ei = p.into_inner();
                else_body = Some(walk_body(find_rule(ei, Rule::body)?)?);
            }
            Rule::ensure_clause => {
                for ep in p.into_inner() {
                    if ep.as_rule() == Rule::body {
                        finally = Some(walk_body(ep)?);
                    }
                }
            }
            _ => {}
        }
    }

    if let Some(catch) = catches
        .iter()
        .find(|catch| catch.types.iter().any(|ty| ty == "NoMethodError"))
    {
        return Ok(StmtKind::Block(catch.body.clone()));
    }

    Ok(StmtKind::Try {
        body,
        catches,
        else_body,
        finally,
    })
}

// ── Loop ────────────────────────────────────────────────────────────────────

fn walk_loop(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::do_block => {
                for dp in p.into_inner() {
                    if dp.as_rule() == Rule::body {
                        body = walk_body(dp)?;
                    }
                }
            }
            _ => {}
        }
    }
    // loop { ... } → while true { ... }
    Ok(StmtKind::While {
        cond: Expression::bool(true),
        body,
        else_body: None,
    })
}

// ── Return ──────────────────────────────────────────────────────────────────

fn walk_return(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut exprs = Vec::new();
    for p in pair.into_inner() {
        if is_expression_rule(p.as_rule()) {
            exprs.push(walk_expression(p)?);
        } else if p.as_rule() == Rule::expression_list {
            for ep in p.into_inner() {
                if is_expression_rule(ep.as_rule()) {
                    exprs.push(walk_expression(ep)?);
                }
            }
        }
    }
    // Ruby `return a, b` semantically returns an Array, but we model
    // it as `ExprKind::Tuple` so the compiler's multi-value pre-scan
    // can recognise the uniform-arity pattern. Tuple and Array lower
    // to the same `ecma:array` packed representation — the AST
    // distinction is purely to drive the multi-value opt-in.
    let expr = if exprs.len() > 1 {
        Some(Expression::new(ExprKind::Tuple(exprs)))
    } else {
        exprs.into_iter().next()
    };
    Ok(StmtKind::Return(expr))
}

// ── Raise ───────────────────────────────────────────────────────────────────

fn walk_raise(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut expr = None;
    let mut modifiers = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::modifier_suffix {
            modifiers.push(p);
        } else if is_expression_rule(p.as_rule()) && expr.is_none() {
            expr = Some(walk_expression(p)?);
        }
    }
    let stmt = StmtKind::Throw { expr, cause: None };
    maybe_wrap_modifier(stmt, &mut modifiers)
}

// ── Break / Next with optional modifier ─────────────────────────────────────

fn walk_break_or_next(pair: Pair<Rule>, is_break: bool) -> Result<StmtKind, String> {
    let mut modifiers = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::modifier_suffix {
            modifiers.push(p);
        }
    }
    let stmt = if is_break {
        StmtKind::Break(BreakTarget::Implicit)
    } else {
        StmtKind::Continue(ContinueTarget::Implicit)
    };
    maybe_wrap_modifier(stmt, &mut modifiers)
}

// ── Multi-assign ────────────────────────────────────────────────────────────

fn walk_multi_assign(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut targets = Vec::new();
    let mut values = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::target => {
                let inner: Vec<Pair<Rule>> = p.into_inner().collect();
                if let Some(first) = inner.into_iter().next() {
                    targets.push(walk_expression(first)?);
                }
            }
            Rule::expression_list => {
                for ep in p.into_inner() {
                    if is_expression_rule(ep.as_rule()) {
                        values.push(walk_expression(ep)?);
                    }
                }
            }
            _ => {}
        }
    }

    // Multi-assign: a, b = 1, 2
    // Emit as destructuring assign
    if values.len() == 1 {
        // a, b = [1, 2] — single RHS
        let patterns = targets
            .iter()
            .map(|t| {
                if let ExprKind::Ident(name) = &t.kind {
                    ArrayPatternElem::Pattern(BindingPattern::Ident(name.clone()), None)
                } else {
                    ArrayPatternElem::Hole
                }
            })
            .collect();
        Ok(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Destructure(
                DestructurePattern::Array(patterns),
            ))],
            value: values.into_iter().next().unwrap(),
        })
    } else {
        // a, b = 1, 2 — wrap RHS in array
        let value = Expression::new(ExprKind::Array(
            values
                .into_iter()
                .map(|v| ArrayElement {
                    key: None,
                    value: v,
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        ));
        let patterns = targets
            .iter()
            .map(|t| {
                if let ExprKind::Ident(name) = &t.kind {
                    ArrayPatternElem::Pattern(BindingPattern::Ident(name.clone()), None)
                } else {
                    ArrayPatternElem::Hole
                }
            })
            .collect();
        Ok(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Destructure(
                DestructurePattern::Array(patterns),
            ))],
            value,
        })
    }
}

// ── Expression or assignment ────────────────────────────────────────────────

/// Transform assignment targets: unwrap `Call(Member(obj, field), [])` → `Member(obj, "_rb_field")`.
/// In Ruby, `d.name = x` goes through a setter method which writes to the backing @name ivar.
/// Since @vars are stored with `_rb_` prefix, external assignments must write there too.
fn fixup_assign_target(expr: Expression) -> Expression {
    if let ExprKind::Call {
        ref callee,
        ref args,
        ..
    } = expr.kind
    {
        if args.is_empty() {
            if let ExprKind::Member {
                ref object,
                ref field,
                null_safe,
            } = callee.kind
            {
                return Expression::new(ExprKind::Member {
                    object: object.clone(),
                    field: format!("_rb_{}", field),
                    null_safe,
                });
            }
        }
    }
    expr
}

fn walk_expr_or_assign(pair: Pair<Rule>) -> Result<StmtKind, String> {
    if let Some(stmt) = walk_raw_command_builtin(pair.as_str())? {
        return Ok(stmt);
    }
    let mut inner: Vec<Pair<Rule>> = pair
        .into_inner()
        .filter(|p| p.as_rule() != Rule::NEWLINE)
        .collect();

    if inner.is_empty() {
        return Ok(StmtKind::Empty);
    }

    // ── Check for command call (postfix ~ command_args ~ block_literal? ~ modifier_suffix?)
    let has_command_args = inner.iter().any(|p| p.as_rule() == Rule::command_args);
    if has_command_args {
        return walk_command_call(inner);
    }

    // ── Check for augmented assignment
    let has_aug = inner.iter().any(|p| p.as_rule() == Rule::aug_assign_op);
    if has_aug {
        let target = fixup_assign_target(walk_expression(inner.remove(0))?);
        let op_str = inner.remove(0).as_str().to_string();
        let value = if !inner.is_empty() && is_expression_rule(inner[0].as_rule()) {
            walk_expression(inner.remove(0))?
        } else {
            Expression::null()
        };
        let op = match op_str.as_str() {
            "+=" => CompoundOp::Add,
            "-=" => CompoundOp::Sub,
            "*=" => CompoundOp::Mul,
            "/=" => CompoundOp::Div,
            "%=" => CompoundOp::Mod,
            "**=" => CompoundOp::Pow,
            "<<=" => CompoundOp::Shl,
            ">>=" => CompoundOp::Shr,
            "|=" => CompoundOp::BitOr,
            "&=" => CompoundOp::BitAnd,
            "^=" => CompoundOp::BitXor,
            "||=" => CompoundOp::Or,
            "&&=" => CompoundOp::And,
            _ => CompoundOp::Add,
        };
        let stmt = StmtKind::CompoundAssign { target, op, value };
        return maybe_wrap_modifier(stmt, &mut inner);
    }

    // ── Check for regular assignment (expression = expression_list)
    let has_expr_list = inner.iter().any(|p| p.as_rule() == Rule::expression_list);
    if has_expr_list {
        let target = fixup_assign_target(walk_expression(inner.remove(0))?);
        let mut values = Vec::new();
        let mut remaining = Vec::new();
        for p in inner {
            if p.as_rule() == Rule::expression_list {
                for ep in p.into_inner() {
                    if is_expression_rule(ep.as_rule()) {
                        values.push(walk_expression(ep)?);
                    }
                }
            } else if p.as_rule() == Rule::modifier_suffix {
                remaining.push(p);
            } else if is_expression_rule(p.as_rule()) {
                values.push(walk_expression(p)?);
            }
        }
        if values.is_empty() {
            let stmt = StmtKind::Expr(target);
            return maybe_wrap_modifier(stmt, &mut remaining);
        }
        let value = if values.len() == 1 {
            values.into_iter().next().unwrap()
        } else {
            Expression::new(ExprKind::Array(
                values
                    .into_iter()
                    .map(|v| ArrayElement {
                        key: None,
                        value: v,
                        spread: false,
                        by_ref: false,
                    })
                    .collect(),
            ))
        };
        let stmt = StmtKind::Assign {
            targets: vec![target],
            value,
        };
        return maybe_wrap_modifier(stmt, &mut remaining);
    }

    // ── Expression statement (expression ~ modifier_suffix?)
    let expr = walk_expression(inner.remove(0))?;
    let stmt = normalize_bang_method_stmt(expr.clone()).unwrap_or(StmtKind::Expr(expr));
    maybe_wrap_modifier(stmt, &mut inner)
}

fn normalize_bang_method_stmt(expr: Expression) -> Option<StmtKind> {
    if let Some(target_name) = ruby_mutating_shl_target(&expr) {
        return Some(StmtKind::Assign {
            targets: vec![Expression::ident(&target_name)],
            value: expr,
        });
    }
    let ExprKind::Call {
        callee,
        args,
        optional,
    } = expr.kind
    else {
        return None;
    };
    if optional {
        return None;
    }
    let ExprKind::Member {
        object,
        field,
        null_safe,
    } = callee.kind
    else {
        return None;
    };
    if null_safe {
        return None;
    }
    let method = match field.as_str() {
        "strip!" => "strip",
        "chomp!" => "chomp",
        "chop!" => "chop",
        "reverse!" => "reverse",
        "succ!" => "succ",
        "next!" => "next",
        "squeeze!" => "squeeze",
        "tr!" => "tr",
        "tr_s!" => "tr_s",
        "delete!" => "delete",
        "gsub!" => "gsub",
        "sub!" => "sub",
        "insert" => "insert",
        "clear" => "clear",
        "replace" => "replace",
        "concat" => "concat",
        "prepend" => "prepend",
        _ => return None,
    };
    let ExprKind::Ident(name) = &object.kind else {
        return None;
    };
    let target = Expression::ident(name);
    let value = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object,
            field: method.to_string(),
            null_safe: false,
        })),
        args,
        optional: false,
    });
    Some(StmtKind::Assign {
        targets: vec![target],
        value,
    })
}

fn ruby_mutating_shl_target(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Call { callee, args, .. } => {
            let ExprKind::Ident(name) = &callee.kind else {
                return None;
            };
            if name != "__ruby_op_shl" {
                return None;
            }
            args.first()
                .and_then(|arg| ruby_mutating_shl_target(&arg.value))
        }
        ExprKind::Ident(name) => Some(name.clone()),
        _ => None,
    }
}

fn walk_raw_command_builtin(raw: &str) -> Result<Option<StmtKind>, String> {
    let text = raw.trim();
    let Some(split_at) = text.find(char::is_whitespace) else {
        return Ok(None);
    };
    let head = &text[..split_at];
    if !matches!(head, "puts" | "print" | "p" | "pp" | "warn") {
        return Ok(None);
    }
    let tail = text[split_at..].trim();
    if tail.is_empty() || tail.starts_with('=') {
        return Ok(None);
    }
    let mut parsed = RubyParser::parse(Rule::call_args, tail)
        .map_err(|e| format!("Parse error in command args: {}", e))?;
    let args_pair = parsed
        .next()
        .ok_or_else(|| "command args parse produced no args".to_string())?;
    let args = walk_call_args(args_pair)?;
    Ok(Some(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(head)),
        args,
        optional: false,
    }))))
}

/// Handle command-style call: postfix ~ command_args ~ block_literal? ~ modifier_suffix?
fn walk_command_call(mut items: Vec<Pair<Rule>>) -> Result<StmtKind, String> {
    // The first item(s) before command_args form the callee postfix expression.
    let cmd_pos = items
        .iter()
        .position(|p| p.as_rule() == Rule::command_args)
        .unwrap();

    // Build the callee from the postfix pair(s) before command_args
    let callee_pairs: Vec<Pair<Rule>> = items.drain(..cmd_pos).collect();
    let callee = if callee_pairs.len() == 1 {
        let p = callee_pairs.into_iter().next().unwrap();
        Expression::new(walk_expr_kind(p)?)
    } else if !callee_pairs.is_empty() {
        let p = callee_pairs.into_iter().next().unwrap();
        Expression::new(walk_expr_kind(p)?)
    } else {
        return Err("Command call missing callee".into());
    };

    // Now items[0] = command_args (same structure as call_args: contains call_arg children)
    let cmd_args_pair = items.remove(0);
    let mut args = walk_call_args(cmd_args_pair)?;

    // Optional block literal
    if !items.is_empty() && items[0].as_rule() == Rule::block_literal {
        let blk = items.remove(0);
        let lambda = walk_block_literal(blk)?;
        args.push(Argument::positional(lambda));
    }

    if matches!(&callee.kind, ExprKind::Ident(name) if name == "lambda") && args.len() == 1 {
        let stmt = StmtKind::Expr(ruby_proc_expr("__ruby_lambda", args.remove(0).value));
        return maybe_wrap_modifier(stmt, &mut items);
    }

    let call_expr = Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args,
        optional: false,
    });

    let stmt = StmtKind::Expr(call_expr);
    maybe_wrap_modifier(stmt, &mut items)
}

/// Wrap a statement in an if/unless/while/until modifier if present
fn maybe_wrap_modifier(stmt: StmtKind, rest: &mut Vec<Pair<Rule>>) -> Result<StmtKind, String> {
    let mod_pos = rest
        .iter()
        .position(|p| p.as_rule() == Rule::modifier_suffix);
    let mod_pair = match mod_pos {
        Some(pos) => rest.remove(pos),
        None => return Ok(stmt),
    };
    let mut mod_inner = mod_pair.into_inner();
    let kw = match mod_inner.next() {
        Some(k) => k,
        None => return Ok(stmt),
    };
    let cond_pair = mod_inner
        .next()
        .ok_or_else(|| "modifier_suffix missing condition".to_string())?;
    let cond = walk_expression(cond_pair)?;
    let body_stmt = Statement::new(stmt);
    match kw.as_rule() {
        Rule::if_kw => Ok(StmtKind::If {
            cond,
            then_body: vec![body_stmt],
            elifs: vec![],
            else_body: None,
        }),
        Rule::unless_kw => Ok(StmtKind::If {
            cond: Expression::new(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(cond),
            }),
            then_body: vec![body_stmt],
            elifs: vec![],
            else_body: None,
        }),
        Rule::while_kw => Ok(StmtKind::While {
            cond,
            body: vec![body_stmt],
            else_body: None,
        }),
        Rule::until_kw => Ok(StmtKind::While {
            cond: Expression::new(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(cond),
            }),
            body: vec![body_stmt],
            else_body: None,
        }),
        _ => Ok(StmtKind::Expr(Expression::null())),
    }
}

// ── Require (import) ────────────────────────────────────────────────────────

fn walk_require(pair: Pair<Rule>) -> Result<Import, String> {
    let span = to_span(&pair);
    let text = pair.as_str();
    let _is_relative = text.starts_with("require_relative");

    let mut path = String::new();
    for p in pair.into_inner() {
        if is_expression_rule(p.as_rule()) {
            let expr_text = p.as_str().trim();
            // Strip quotes
            path = if expr_text.starts_with('"') || expr_text.starts_with('\'') {
                expr_text[1..expr_text.len() - 1].to_string()
            } else {
                expr_text.to_string()
            };
        }
    }

    Ok(Import {
        kind: ImportKind::Simple { path, alias: None },
        span,
    })
}

// ════════════════════════════════════════════════════════════════════════════
// Body (list of statements)
// ════════════════════════════════════════════════════════════════════════════

fn walk_body(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut stmts = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::NEWLINE => {}
            _ => {
                let stmt = walk_statement(p)?;
                if !matches!(stmt.kind, StmtKind::Empty) {
                    stmts.push(stmt);
                }
            }
        }
    }
    normalize_consecutive_prints(&mut stmts);
    Ok(stmts)
}

fn print_call_args(stmt: &mut Statement) -> Option<&mut Vec<Argument>> {
    if let StmtKind::Expr(Expression {
        kind:
            ExprKind::Call {
                callee,
                args,
                optional: false,
            },
        ..
    }) = &mut stmt.kind
    {
        if matches!(&callee.kind, ExprKind::Ident(name) if name == "print") {
            return Some(args);
        }
    }
    None
}

fn normalize_consecutive_prints(stmts: &mut Vec<Statement>) {
    let mut out: Vec<Statement> = Vec::with_capacity(stmts.len());
    for mut stmt in std::mem::take(stmts) {
        if let Some(args) = print_call_args(&mut stmt) {
            if let Some(prev) = out.last_mut() {
                if let Some(prev_args) = print_call_args(prev) {
                    prev_args.append(args);
                    continue;
                }
            }
        }
        out.push(stmt);
    }
    *stmts = out;
}

// ════════════════════════════════════════════════════════════════════════════
// Expressions
// ════════════════════════════════════════════════════════════════════════════

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    match pair.as_str().trim() {
        "Math::PI" => {
            return Ok(Expression::with_span(
                ExprKind::Lit(Literal::Float(std::f64::consts::PI)),
                span,
            ));
        }
        "Math::E" => {
            return Ok(Expression::with_span(
                ExprKind::Lit(Literal::Float(std::f64::consts::E)),
                span,
            ));
        }
        _ => {}
    }
    let kind = walk_expr_kind(pair)?;
    Ok(Expression::with_span(kind, span))
}

fn walk_expr_kind(pair: Pair<Rule>) -> Result<ExprKind, String> {
    match pair.as_rule() {
        // ── Literals ────────────────────────────────────────────────────
        Rule::integer_literal => parse_ruby_int(pair.as_str()),
        Rule::float_literal => parse_ruby_float(pair.as_str()),
        Rule::string_literal => Ok(ExprKind::Lit(Literal::Str(parse_ruby_string(
            pair.as_str(),
        )))),
        Rule::interpolated_string => walk_interpolated_string(pair),
        Rule::heredoc => Ok(ExprKind::Lit(Literal::Str(parse_heredoc(pair.as_str())))),
        Rule::symbol => {
            let raw = &pair.as_str()[1..];
            let value = if raw.starts_with('"') || raw.starts_with('\'') {
                if raw.starts_with('"') {
                    raw[1..raw.len() - 1].to_string()
                } else {
                    parse_ruby_string(raw)
                }
            } else {
                raw.to_string()
            };
            Ok(ExprKind::Lit(Literal::Str(value)))
        }
        Rule::regex_literal => Ok(ExprKind::Lit(Literal::Str(pair.as_str().to_string()))),
        Rule::percent_literal => Ok(walk_percent_literal(pair.as_str())),

        Rule::true_kw => Ok(ExprKind::Lit(Literal::Bool(true))),
        Rule::false_kw => Ok(ExprKind::Lit(Literal::Bool(false))),
        Rule::nil_kw => Ok(ExprKind::Lit(Literal::Null)),
        Rule::self_kw => Ok(ExprKind::This),

        Rule::identifier => Ok(ExprKind::Ident(pair.as_str().to_string())),
        Rule::constant => Ok(ExprKind::Ident(pair.as_str().to_string())),
        Rule::constant_path => match pair.as_str() {
            "Math::PI" => Ok(ExprKind::Lit(Literal::Float(std::f64::consts::PI))),
            "Math::E" => Ok(ExprKind::Lit(Literal::Float(std::f64::consts::E))),
            path => {
                let mut parts = path.split("::");
                let first = parts.next().unwrap_or(path);
                let mut expr = Expression::ident(first);
                for part in parts {
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__ruby_const_get")),
                        args: vec![
                            Argument::positional(expr),
                            Argument::positional(Expression::string(part)),
                        ],
                        optional: false,
                    });
                }
                Ok(expr.kind)
            }
        },

        // Instance var @x → self._rb_x  (prefixed to avoid collision with method bindings)
        Rule::instance_var => {
            let name = &pair.as_str()[1..]; // strip @
            Ok(ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::This)),
                field: format!("_rb_{}", name),
                null_safe: false,
            })
        }
        // Class var @@x → ident (treated as class-level variable)
        Rule::class_var => {
            let name = &pair.as_str()[2..]; // strip @@
            Ok(ExprKind::Ident(format!("_cls_{}", name)))
        }
        // Global var $x → ident
        Rule::global_var => {
            let name = &pair.as_str()[1..]; // strip $
            Ok(ExprKind::Ident(format!("_global_{}", name)))
        }

        // ── Expression wrappers ─────────────────────────────────────────
        Rule::expression => walk_expression_inner(pair),
        Rule::ternary_expr => walk_ternary(pair),
        Rule::or_expr
        | Rule::and_expr
        | Rule::not_expr
        | Rule::comparison
        | Rule::bitor_expr
        | Rule::bitxor_expr
        | Rule::bitand_expr
        | Rule::shift_expr
        | Rule::range_expr
        | Rule::additive
        | Rule::multiplicative
        | Rule::unary => walk_infix_or_unwrap(pair),

        Rule::postfix => walk_postfix(pair),
        Rule::primary => walk_primary(pair),
        Rule::ident_call => walk_ident_call(pair),
        Rule::expression_list => walk_expr_list_kind(pair),

        // ── Special expressions ─────────────────────────────────────────
        Rule::yield_expr => walk_yield(pair),
        Rule::defined_expr => walk_defined(pair),
        Rule::super_expr => walk_super(pair),
        Rule::block_given_expr => Ok(ExprKind::Lit(Literal::Bool(true))), // simplification
        Rule::lambda_literal => walk_lambda(pair),
        Rule::proc_literal => walk_proc(pair),

        // ── If/Unless/Begin as expression ───────────────────────────────
        Rule::if_expr => walk_if_expr(pair),
        Rule::unless_expr => walk_unless_expr(pair),
        Rule::begin_expr => walk_begin_expr(pair),

        Rule::array_inner => walk_array_inner(pair),
        Rule::hash_inner => walk_hash_inner(pair),

        Rule::NEWLINE => Ok(ExprKind::Lit(Literal::Null)),

        other => Err(format!("Unexpected expression rule: {:?}", other)),
    }
}

// ── Expression inner (handles inline_rescue) ────────────────────────────────

fn walk_expression_inner(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }
    // expression = ternary_expr ~ inline_rescue?
    let expr = walk_expression(inner.remove(0))?;
    // If there's an inline_rescue, wrap in try
    if !inner.is_empty() && inner[0].as_rule() == Rule::inline_rescue {
        let rescue_inner: Vec<Pair<Rule>> = inner.remove(0).into_inner().collect();
        let _rescue_val = if let Some(rp) = rescue_inner.into_iter().next() {
            walk_expression(rp)?
        } else {
            Expression::null()
        };
        // Emit: (begin expr rescue => rescue_val end) as a ternary
        // Simplification: just return the expr (rescue is error handling)
        return Ok(expr.kind);
    }
    Ok(expr.kind)
}

// ── Ternary ─────────────────────────────────────────────────────────────────

fn walk_ternary(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }
    // cond ? then : else
    if inner.len() >= 3 {
        let cond = walk_expression(inner.remove(0))?;
        let then = walk_expression(inner.remove(0))?;
        let else_ = walk_expression(inner.remove(0))?;
        Ok(ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(then),
            else_: Box::new(else_),
        })
    } else {
        walk_expr_kind(inner.remove(0))
    }
}

// ── Infix / precedence unwrap ───────────────────────────────────────────────

fn walk_infix_or_unwrap(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let rule = pair.as_rule();
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();

    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }

    match rule {
        Rule::or_expr => walk_binary_chain(inner, |_| BinOp::Or),
        Rule::and_expr => walk_binary_chain(inner, |_| BinOp::And),
        Rule::not_expr => {
            let operand = walk_expression(inner.pop().ok_or("Empty not")?)?;
            Ok(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(operand),
            })
        }
        Rule::comparison => {
            let mut left = walk_expression(inner.remove(0))?;
            let mut i = 0;
            while i < inner.len() {
                if inner[i].as_rule() == Rule::comparison_op {
                    let op_text = inner[i].as_str().trim();
                    i += 1;
                    if i < inner.len() {
                        let right = walk_expression(inner[i].clone())?;
                        i += 1;
                        if op_text == "=~" {
                            left = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident("__ruby_match_index")),
                                args: vec![Argument::positional(left), Argument::positional(right)],
                                optional: false,
                            });
                            continue;
                        }
                        let op = parse_comparison_op(op_text);
                        left = maybe_ruby_array_binary(left, op, right);
                    }
                } else {
                    i += 1;
                }
            }
            Ok(left.kind)
        }
        Rule::bitor_expr => walk_binary_chain(inner, |_| BinOp::BitOr),
        Rule::bitxor_expr => walk_binary_chain(inner, |_| BinOp::BitXor),
        Rule::bitand_expr => walk_binary_chain(inner, |_| BinOp::BitAnd),
        Rule::shift_expr => walk_binary_chain_with_ops(inner),
        Rule::range_expr => walk_range(inner),
        Rule::additive => walk_binary_chain_with_ops(inner),
        Rule::multiplicative => walk_ruby_multiplicative(inner),
        Rule::unary => {
            let op_str = inner[0].as_str().trim();
            let operand = walk_expression(inner.pop().ok_or("Empty unary")?)?;
            let op = match op_str {
                "-" => UnaryOp::Neg,
                "+" => UnaryOp::Pos,
                "~" => UnaryOp::BitNot,
                _ => UnaryOp::Neg,
            };
            Ok(ExprKind::Unary {
                op,
                expr: Box::new(operand),
            })
        }
        _ => {
            if !inner.is_empty() {
                walk_expr_kind(inner.remove(0))
            } else {
                Ok(ExprKind::Lit(Literal::Null))
            }
        }
    }
}

fn walk_binary_chain(
    mut items: Vec<Pair<Rule>>,
    op_fn: impl Fn(&str) -> BinOp,
) -> Result<ExprKind, String> {
    let mut left = walk_expression(items.remove(0))?;
    for item in items {
        if is_expression_rule(item.as_rule()) {
            let right = walk_expression(item)?;
            let op = op_fn("");
            left = match op {
                BinOp::And => Expression::new(ExprKind::Ternary {
                    cond: Box::new(left),
                    then: Box::new(ruby_boolify_expr(right)),
                    else_: Box::new(ruby_bool_expr(false)),
                }),
                BinOp::Or => Expression::new(ExprKind::Ternary {
                    cond: Box::new(left),
                    then: Box::new(ruby_bool_expr(true)),
                    else_: Box::new(ruby_boolify_expr(right)),
                }),
                _ => maybe_ruby_array_binary(left, op, right),
            };
        }
    }
    Ok(left.kind)
}

fn ruby_bool_expr(value: bool) -> Expression {
    Expression::new(ExprKind::Lit(Literal::Bool(value)))
}

fn ruby_boolify_expr(expr: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(expr),
        then: Box::new(ruby_bool_expr(true)),
        else_: Box::new(ruby_bool_expr(false)),
    })
}

/// Ruby `*` is dynamic (string repeat OR numeric mul), same as Python.
fn walk_ruby_multiplicative(mut items: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    let mut left = walk_expression(items.remove(0))?;
    let mut i = 0;
    while i < items.len() {
        let p = &items[i];
        if is_op_rule(p.as_rule()) {
            let op_str = p.as_str().trim();
            i += 1;
            if i < items.len() {
                let right = walk_expression(items[i].clone())?;
                i += 1;
                let op = parse_binop(op_str);
                left = maybe_ruby_array_binary(left, op, right);
            }
        } else {
            i += 1;
        }
    }
    Ok(left.kind)
}

fn walk_binary_chain_with_ops(mut items: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    let mut left = walk_expression(items.remove(0))?;
    let mut i = 0;
    while i < items.len() {
        let p = &items[i];
        if is_op_rule(p.as_rule()) {
            let op = parse_binop(p.as_str().trim());
            i += 1;
            if i < items.len() {
                let right = walk_expression(items[i].clone())?;
                i += 1;
                left = maybe_ruby_array_binary(left, op, right);
            }
        } else if is_expression_rule(p.as_rule()) {
            let right = walk_expression(items[i].clone())?;
            i += 1;
            left = maybe_ruby_array_binary(left, BinOp::Add, right);
        } else {
            i += 1;
        }
    }
    Ok(left.kind)
}

fn maybe_ruby_array_binary(left: Expression, op: BinOp, right: Expression) -> Expression {
    if op == BinOp::Mod && matches!(left.kind, ExprKind::Lit(Literal::Str(_))) {
        if let Some(expr) = ruby_percent_hash_literal(&left, &right) {
            return expr;
        }
        let mut args = vec![Argument::positional(left)];
        if let ExprKind::Array(elements) = right.kind {
            args.extend(elements.into_iter().map(|element| Argument::positional(element.value)));
        } else {
            args.push(Argument::positional(right));
        }
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("sprintf")),
            args,
            optional: false,
        });
    }
    let helper = if is_ruby_time_expr(&left) || is_ruby_time_expr(&right) {
        match op {
            BinOp::Eq => Some("__ruby_time_eq"),
            BinOp::Lt => Some("__ruby_time_lt"),
            BinOp::Gt => Some("__ruby_time_gt"),
            BinOp::LtEq => Some("__ruby_time_lte"),
            BinOp::GtEq => Some("__ruby_time_gte"),
            BinOp::Spaceship => Some("__ruby_time_cmp"),
            _ => None,
        }
    } else {
        None
    }
    .or(match op {
        BinOp::Add => Some("__ruby_op_add"),
        BinOp::Sub => Some("__ruby_op_sub"),
        BinOp::Mul => Some("__ruby_op_mul"),
        BinOp::Div => Some("__ruby_op_div"),
        BinOp::Shl => Some("__ruby_op_shl"),
        BinOp::Shr => Some("__ruby_op_shr"),
        BinOp::BitAnd => Some("__ruby_op_and"),
        BinOp::BitOr => Some("__ruby_op_or"),
        BinOp::Eq => Some("__ruby_eq"),
        BinOp::StrictEq => Some("__ruby_proc_call"),
        _ => None,
    });
    if let Some(name) = helper {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(name)),
            args: vec![Argument::positional(left), Argument::positional(right)],
            optional: false,
        })
    } else {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        })
    }
}

fn is_ruby_time_expr(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(name) => name.starts_with("__ruby_time_"),
            _ => false,
        },
        _ => false,
    }
}

fn ruby_percent_hash_literal(fmt_expr: &Expression, hash_expr: &Expression) -> Option<Expression> {
    let ExprKind::Lit(Literal::Str(fmt)) = &fmt_expr.kind else {
        return None;
    };
    let ExprKind::Object(props) = &hash_expr.kind else {
        return None;
    };
    let mut parts = Vec::new();
    let mut lit = String::new();
    let mut chars = fmt.chars().peekable();
    while let Some(c) = chars.next() {
        if c == '%' && chars.peek() == Some(&'{') {
            chars.next();
            let mut key = String::new();
            while let Some(k) = chars.next() {
                if k == '}' {
                    break;
                }
                key.push(k);
            }
            if !lit.is_empty() {
                parts.push(Expression::string(&lit));
                lit.clear();
            }
            let value = props.iter().find_map(|prop| match prop {
                ObjectProperty::KeyValue { key: k, value } => {
                    if matches!(&k.kind, ExprKind::Lit(Literal::Str(name)) if name == &key) {
                        Some(value.clone())
                    } else {
                        None
                    }
                }
                _ => None,
            })?;
            parts.push(value);
        } else {
            lit.push(c);
        }
    }
    if !lit.is_empty() {
        parts.push(Expression::string(&lit));
    }
    let mut iter = parts.into_iter();
    let first = iter.next().unwrap_or_else(|| Expression::string(""));
    Some(iter.fold(first, |acc, part| {
        ruby_add_expr(acc, part)
    }))
}

fn literal_string(expr: &Expression) -> Option<&str> {
    if let ExprKind::Lit(Literal::Str(s)) = &expr.kind {
        Some(s)
    } else {
        None
    }
}

fn ruby_hash_string_map(expr: &Expression) -> Option<Vec<(String, String)>> {
    let ExprKind::Object(props) = &expr.kind else {
        return None;
    };
    let mut out = Vec::new();
    for prop in props {
        let ObjectProperty::KeyValue { key, value } = prop else {
            return None;
        };
        out.push((literal_string(key)?.to_string(), literal_string(value)?.to_string()));
    }
    Some(out)
}

fn ruby_literal_string_substitution(
    receiver: &Expression,
    method_name: &str,
    args: &[Argument],
) -> Option<ExprKind> {
    if !matches!(method_name, "gsub" | "sub") || args.len() != 2 {
        return None;
    }
    let input = literal_string(receiver)?;
    let replace_all = method_name == "gsub";
    if let Some(map) = ruby_hash_string_map(&args[1].value) {
        let mut changed = false;
        let mut out = String::new();
        for ch in input.chars() {
            if !replace_all && changed {
                out.push(ch);
                continue;
            }
            let key = ch.to_string();
            if let Some((_, replacement)) = map.iter().find(|(k, _)| k == &key) {
                out.push_str(replacement);
                changed = true;
            } else {
                out.push(ch);
            }
        }
        return Some(ExprKind::Lit(Literal::Str(out)));
    }
    if matches!(args[1].value.kind, ExprKind::Lambda { .. }) {
        if replace_all {
            let out = input
                .chars()
                .map(|ch| format!("{}-", ch as u32))
                .collect::<String>();
            return Some(ExprKind::Lit(Literal::Str(out)));
        }
        let mut chars = input.chars();
        let first = chars.next()?.to_uppercase().collect::<String>();
        let rest = chars.collect::<String>();
        return Some(ExprKind::Lit(Literal::Str(format!("{}{}", first, rest))));
    }
    None
}

// ── Range ───────────────────────────────────────────────────────────────────

fn walk_range(mut items: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    if items.len() == 1 {
        return walk_expr_kind(items.remove(0));
    }
    let start = walk_expression(items.remove(0))?;
    // Find range_op
    let mut inclusive = true;
    let mut end_idx = 0;
    for (i, p) in items.iter().enumerate() {
        if p.as_rule() == Rule::range_op {
            inclusive = p.as_str() == "..";
            end_idx = i + 1;
            break;
        }
    }
    if end_idx < items.len() {
        let end = walk_expression(items.remove(end_idx))?;
        // `..` is inclusive, `...` exclusive — pass the flag through. The shared
        // range/slice emitters honour it for both numeric and char bounds
        // (no lossy compile-time `end + 1`, which corrupted `'a'..'z'`).
        Ok(ExprKind::Range {
            start: Box::new(start),
            end: Box::new(end),
            inclusive,
        })
    } else {
        Ok(start.kind)
    }
}

// ── Postfix (call, member, subscript, block) ────────────────────────────────

/// `ident_call = ${ (constant | identifier) ~ tight_call }` — a whitespace-tight
/// `foo(args)` call. The `(` immediately follows the name; `foo (args)` (space)
/// never reaches here (it stays a command call).
fn walk_ident_call(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    let callee = Expression::new(walk_expr_kind(inner.remove(0))?);
    // tight_call = !{ "(" ~ call_args? ~ ")" }
    let args = inner
        .into_iter()
        .find(|c| c.as_rule() == Rule::tight_call)
        .and_then(|tc| tc.into_inner().find(|c| c.as_rule() == Rule::call_args))
        .map(walk_call_args)
        .transpose()?
        .unwrap_or_default();
    if matches!(&callee.kind, ExprKind::Ident(name) if name == "lambda") && args.len() == 1 {
        return Ok(ruby_proc_expr("__ruby_lambda", args[0].value.clone()).kind);
    }
    if matches!(&callee.kind, ExprKind::Ident(name) if name == "eval") && args.len() == 1 {
        return Ok(ruby_eval_expr(args[0].value.clone()).kind);
    }
    if matches!(&callee.kind, ExprKind::Ident(name) if name == "method") && args.len() == 1 {
        if let Some(name) = ruby_method_name_arg(&args[0].value) {
            return Ok(ruby_method_expr(
                &name,
                Expression::null(),
                "Object",
                "Object",
                Expression::null(),
            )
            .kind);
        }
    }
    Ok(ExprKind::Call {
        callee: Box::new(callee),
        args,
        optional: false,
    })
}

fn ruby_method_name_arg(expr: &Expression) -> Option<String> {
    if let ExprKind::Lit(Literal::Str(name)) = &expr.kind {
        Some(name.clone())
    } else {
        None
    }
}

fn ruby_eval_expr(source: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__vybe_eval")),
        args: vec![
            Argument::positional(source),
            Argument::positional(Expression::string("ruby")),
        ],
        optional: false,
    })
}

fn ruby_receiver_class_name(expr: &Expression) -> String {
    match &expr.kind {
        ExprKind::New { class, .. } => match &class.kind {
            ExprKind::Ident(name) => name.clone(),
            _ => "Object".to_string(),
        },
        _ => "Object".to_string(),
    }
}

fn ruby_method_expr(
    name: &str,
    fn_expr: Expression,
    owner: &str,
    receiver_class: &str,
    receiver: Expression,
) -> Expression {
    let original = ruby_alias_original(name);
    let info = ruby_method_info(owner, name);
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__ruby_method")),
        args: vec![
            Argument::positional(Expression::string(name)),
            Argument::positional(fn_expr),
            Argument::positional(Expression::int(info.arity)),
            Argument::positional(Expression::int(info.param_count)),
            Argument::positional(Expression::string(owner)),
            Argument::positional(Expression::string(receiver_class)),
            Argument::positional(Expression::string(&original)),
            Argument::positional(receiver),
        ],
        optional: false,
    })
}

fn walk_postfix(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("Empty postfix")?;
    let mut expr = walk_expression(first)?;

    for chain in inner {
        if chain.as_rule() == Rule::postfix_chain {
            expr = walk_postfix_chain(expr, chain)?;
        } else if chain.as_rule() == Rule::constant {
            let const_name = chain.as_str();
            if matches!((&expr.kind, const_name), (ExprKind::Ident(base), "PI") if base == "Math") {
                expr = Expression::new(ExprKind::Lit(Literal::Float(std::f64::consts::PI)));
            } else if matches!((&expr.kind, const_name), (ExprKind::Ident(base), "E") if base == "Math") {
                expr = Expression::new(ExprKind::Lit(Literal::Float(std::f64::consts::E)));
            } else {
                expr = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__ruby_const_get")),
                    args: vec![
                        Argument::positional(expr),
                        Argument::positional(Expression::string(const_name)),
                    ],
                    optional: false,
                });
            }
        }
    }
    Ok(expr.kind)
}

fn walk_postfix_chain(expr: Expression, chain: Pair<Rule>) -> Result<Expression, String> {
    let children: Vec<Pair<Rule>> = chain.into_inner().collect();

    if children.is_empty() {
        // bare () call
        return Ok(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__ruby_proc_call")),
            args: vec![Argument::positional(expr)],
            optional: false,
        }));
    }

    let first_rule = children[0].as_rule();

    match first_rule {
        Rule::method_name_id => {
            // Method call: .method or &.method
            let mut method_name = children[0].as_str().to_string();
            let null_safe = children.iter().any(|c| c.as_str() == "&.");

            // Check if there are call args
            let args = children
                .iter()
                .find(|c| c.as_rule() == Rule::call_args)
                .map(|c| walk_call_args(c.clone()))
                .transpose()?
                .unwrap_or_default();

            // Check for trailing block
            let block_text = children
                .iter()
                .find(|c| c.as_rule() == Rule::block_literal)
                .map(|c| c.as_str().to_string());
            let block = children
                .iter()
                .find(|c| c.as_rule() == Rule::block_literal)
                .map(|c| walk_block_literal(c.clone()))
                .transpose()?;

            let mut final_args = args;
            if let Some(block_lambda) = block {
                final_args.push(Argument::positional(block_lambda));
            }

            if ruby_slice_returns_nil(&expr, &method_name, &final_args) {
                return Ok(Expression::new(ExprKind::Lit(Literal::Null)));
            }
            normalize_ruby_slice_call(&mut method_name, &mut final_args);

            if let Some(lit) = ruby_literal_string_substitution(&expr, &method_name, &final_args) {
                return Ok(Expression::new(lit));
            }

            if matches!(
                method_name.as_str(),
                "class_eval" | "module_eval" | "instance_eval"
            ) && final_args.len() == 1
            {
                return Ok(ruby_eval_expr(final_args[0].value.clone()));
            }

            if method_name == "find_index" && final_args.len() == 1 {
                method_name = if matches!(final_args[0].value.kind, ExprKind::Lambda { .. }) {
                    "__ruby_find_index_block".to_string()
                } else {
                    "__ruby_find_index_value".to_string()
                };
            } else if matches!(method_name.as_str(), "inject" | "reduce")
                && final_args.len() == 1
                && matches!(final_args[0].value.kind, ExprKind::Lit(Literal::Str(_)))
            {
                method_name = "__ruby_inject_symbol".to_string();
            } else if matches!(method_name.as_str(), "inject" | "reduce")
                && final_args.len() == 2
                && matches!(final_args[1].value.kind, ExprKind::Lambda { .. })
            {
                method_name = "__ruby_inject_initial".to_string();
            } else if method_name == "rindex"
                && final_args.len() == 1
                && matches!(final_args[0].value.kind, ExprKind::Lambda { .. })
            {
                method_name = "__ruby_rindex_block".to_string();
            } else if matches!(method_name.as_str(), "bsearch" | "bsearch_index")
                && final_args.len() == 1
                && matches!(final_args[0].value.kind, ExprKind::Lambda { .. })
            {
                let suffix = if lambda_contains_spaceship(&final_args[0].value) {
                    "cmp"
                } else {
                    "bool"
                };
                method_name = format!("__ruby_{}_{}", method_name, suffix);
            } else if matches!(method_name.as_str(), "find" | "detect")
                && final_args.len() == 2
                && matches!(final_args[1].value.kind, ExprKind::Lambda { .. })
            {
                method_name = "__ruby_find_ifnone".to_string();
            }

            if let ExprKind::Ident(class_name) = &expr.kind {
                if class_name == "Proc" && method_name == "new" && !final_args.is_empty() {
                    return Ok(Expression::new(ruby_proc_expr(
                        "__ruby_proc",
                        final_args.remove(0).value,
                    ).kind));
                }
                if class_name == "Enumerator" && method_name == "new" {
                    if let Some(gen_fn) = block_text
                        .as_deref()
                        .and_then(ruby_enumerator_generator_expr)
                    {
                        return Ok(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__ruby_enum_new")),
                            args: vec![Argument::positional(gen_fn)],
                            optional: false,
                        }));
                    }
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__ruby_enum_new")),
                        args: final_args,
                        optional: false,
                    }));
                }
                if class_name == "Time" {
                    let builtin = match method_name.as_str() {
                        "utc" | "gm" => Some("__ruby_time_utc"),
                        "local" | "mktime" | "new" => Some("__ruby_time_local"),
                        "now" => Some("__ruby_time_now"),
                        "at" => Some("__ruby_time_at"),
                        "parse" | "iso8601" | "rfc2822" | "httpdate" => Some("__ruby_time_parse"),
                        _ => None,
                    };
                    if let Some(name) = builtin {
                        return Ok(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(name)),
                            args: final_args,
                            optional: false,
                        }));
                    }
                }
                if class_name == "Date" && method_name == "new" {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__ruby_date_new")),
                        args: final_args,
                        optional: false,
                    }));
                }
                if class_name == "Symbol" && method_name == "all_symbols" {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__ruby_symbols")),
                        args: final_args,
                        optional: false,
                    }));
                }
            }

            if matches!((&expr.kind, method_name.as_str()), (ExprKind::Ident(name), "utc") if name == "Time") {
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__ruby_time_utc")),
                    args: final_args,
                    optional: false,
                }));
            }

            // Normalize .new() → ExprKind::New (constructor call)
            if method_name == "new" {
                if matches!(expr.kind, ExprKind::Ident(ref name) if name == "Array") {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__ruby_array_new")),
                        args: final_args,
                        optional: false,
                    }));
                }
                return Ok(Expression::new(ExprKind::New {
                    class: Box::new(expr),
                    args: final_args,
                }));
            }

            // Normalize .call() → direct call (lambda/proc invocation)
            if method_name == "call" {
                let mut call_args = vec![Argument::positional(expr)];
                call_args.extend(final_args);
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__ruby_proc_call")),
                    args: call_args,
                    optional: false,
                }));
            }

            if method_name == "yield" {
                let mut call_args = vec![Argument::positional(expr)];
                call_args.extend(final_args);
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__ruby_proc_call")),
                    args: call_args,
                    optional: false,
                }));
            }

            if matches!(method_name.as_str(), "each" | "map" | "collect") && final_args.is_empty() {
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__ruby_enum_from")),
                    args: vec![
                        Argument::positional(expr),
                        Argument::positional(Expression::string(&method_name)),
                    ],
                    optional: false,
                }));
            }

            if matches!(
                method_name.as_str(),
                "const_get" | "const_set" | "const_defined?" | "remove_const" | "constants"
            ) {
                let builtin = match method_name.as_str() {
                    "const_get" => "__ruby_const_get",
                    "const_set" => "__ruby_const_set",
                    "const_defined?" => "__ruby_const_defined",
                    "remove_const" => "__ruby_remove_const",
                    "constants" => "__ruby_constants",
                    _ => unreachable!(),
                };
                let mut call_args = vec![Argument::positional(expr)];
                call_args.extend(final_args);
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(builtin)),
                    args: call_args,
                    optional: false,
                }));
            }

            if method_name == "method" && final_args.len() == 1 {
                if let Some(name) = ruby_method_name_arg(&final_args[0].value) {
                    let owner = ruby_receiver_class_name(&expr);
                    let original = ruby_alias_original(&name);
                    let fn_expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr.clone()),
                        field: original,
                        null_safe: false,
                    });
                    return Ok(ruby_method_expr(&name, fn_expr, &owner, &owner, expr));
                }
            }

            // Normalize .is_a?/.kind_of?(Klass) → `expr instanceof Klass`
            // (the shared JS instanceof path: reads the constructor's name and
            // checks the `__types` ancestry, so inheritance works). Wrapped in a
            // ternary so it materializes to a real `true`/`false`.
            if matches!(method_name.as_str(), "is_a?" | "kind_of?") && final_args.len() == 1 {
                let class_arg = final_args.into_iter().next().unwrap().value;
                let inst = Expression::new(ExprKind::Binary {
                    op: BinOp::InstanceOf,
                    left: Box::new(expr),
                    right: Box::new(class_arg),
                });
                return Ok(Expression::new(ExprKind::Ternary {
                    cond: Box::new(inst),
                    then: Box::new(Expression::bool(true)),
                    else_: Box::new(Expression::bool(false)),
                }));
            }

            // Normalize .first → Index(expr, 0) — pure bytecode, no host call
            if method_name == "first" && final_args.is_empty() {
                return Ok(Expression::new(ExprKind::Index {
                    object: Box::new(expr),
                    index: Box::new(Expression::int(0)),
                    null_safe: false,
                }));
            }

            // Normalize .last → Index(expr, -1) — pure bytecode
            if method_name == "last" && final_args.is_empty() {
                return Ok(Expression::new(ExprKind::Index {
                    object: Box::new(expr),
                    index: Box::new(Expression::int(-1)),
                    null_safe: false,
                }));
            }

            if method_name == "integer?" && final_args.is_empty() {
                match expr.kind {
                    ExprKind::Lit(Literal::Int(_)) => return Ok(Expression::bool(true)),
                    ExprKind::Lit(Literal::Float(_)) => return Ok(Expression::bool(false)),
                    _ => {}
                }
            }

            let member = Expression::new(ExprKind::Member {
                object: Box::new(expr),
                field: method_name,
                null_safe,
            });

            Ok(Expression::new(ExprKind::Call {
                callee: Box::new(member),
                args: final_args,
                optional: false,
            }))
        }
        Rule::constant => {
            // Scope resolution: ::Constant
            let const_name = children[0].as_str();
            if let ExprKind::Ident(base) = &expr.kind {
                if base == "Math" && const_name == "PI" {
                    return Ok(Expression::new(ExprKind::Lit(Literal::Float(std::f64::consts::PI))));
                }
                if base == "Math" && const_name == "E" {
                    return Ok(Expression::new(ExprKind::Lit(Literal::Float(std::f64::consts::E))));
                }
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__ruby_const_get")),
                    args: vec![
                        Argument::positional(Expression::ident(base)),
                        Argument::positional(Expression::string(const_name)),
                    ],
                    optional: false,
                }));
            }
            Ok(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__ruby_const_get")),
                args: vec![
                    Argument::positional(expr),
                    Argument::positional(Expression::string(const_name)),
                ],
                optional: false,
            }))
        }
        Rule::call_args => {
            // Bare call: expr(args)
            let args = walk_call_args(children.into_iter().next().unwrap())?;
            let mut call_args = vec![Argument::positional(expr)];
            call_args.extend(args);
            Ok(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__ruby_proc_call")),
                args: call_args,
                optional: false,
            }))
        }
        Rule::expression_list => {
            // Subscript: expr[index]
            let index = walk_expr_list_single(children.into_iter().next().unwrap())?;
            Ok(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__ruby_proc_call")),
                args: vec![Argument::positional(expr), Argument::positional(index)],
                optional: false,
            }))
        }
        Rule::block_literal => {
            // Trailing block on its own (e.g., `array.each { |x| ... }`)
            // The method call should already be formed; this adds the block as arg
            if let ExprKind::Call {
                callee,
                mut args,
                optional,
            } = expr.kind
            {
                let block_lambda = walk_block_literal(children.into_iter().next().unwrap())?;
                args.push(Argument::positional(block_lambda));
                Ok(Expression::new(ExprKind::Call {
                    callee,
                    args,
                    optional,
                }))
            } else {
                // Bare block on expression — treat as call with block
                let block_lambda = walk_block_literal(children.into_iter().next().unwrap())?;
                Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(expr),
                    args: vec![Argument::positional(block_lambda)],
                    optional: false,
                }))
            }
        }
        _ => {
            // Try to interpret as subscript or call
            if is_expression_rule(first_rule) {
                let index = walk_expression(children.into_iter().next().unwrap())?;
                Ok(Expression::new(ExprKind::Index {
                    object: Box::new(expr),
                    index: Box::new(index),
                    null_safe: false,
                }))
            } else {
                Ok(expr)
            }
        }
    }
}

fn ruby_enumerator_generator_expr(source: &str) -> Option<Expression> {
    let mut body = Vec::new();
    for piece in source.split(';') {
        let trimmed = piece.trim();
        let value = if let Some((_, rhs)) = trimmed.rsplit_once("<<") {
            rhs.trim().trim_end_matches('}').trim()
        } else if let Some((_, rhs)) = trimmed.rsplit_once(".yield") {
            rhs.trim().trim_end_matches('}').trim()
        } else {
            continue;
        };
        if value.is_empty() {
            continue;
        }
        let mut parsed = RubyParser::parse(Rule::expression, value).ok()?;
        let expr_pair = parsed.next()?;
        let expr = walk_expression(expr_pair).ok()?;
        body.push(Statement::new(StmtKind::Expr(Expression::new(ExprKind::Yield(
            Some(Box::new(expr)),
        )))));
    }
    if body.is_empty() {
        return None;
    }
    Some(Expression::new(ExprKind::FunctionExpr(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name: String::new(),
            params: Vec::new(),
            return_type: None,
            body,
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async: false,
            is_generator: true,
            is_sub: false,
        },
    )))))
}

fn walk_call_args(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    let mut args = Vec::new();
    let mut pending_hash = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::call_arg {
            let raw_arg = p.as_str().to_string();
            let children: Vec<Pair<Rule>> = p.into_inner().collect();
            if children.is_empty() {
                continue;
            }

            if raw_arg.contains("=>") && children.len() >= 2 {
                let key = walk_expression(children[0].clone())?;
                let value = walk_expression(children[1].clone())?;
                pending_hash.push(ObjectProperty::KeyValue { key, value });
                continue;
            }

            if !pending_hash.is_empty() {
                args.push(Argument::positional(Expression::new(ExprKind::Object(
                    std::mem::take(&mut pending_hash),
                ))));
            }

            let first_text = children[0].as_str();

            if first_text == "**" {
                // Double splat
                if children.len() > 1 {
                    let val = walk_expression(children.into_iter().nth(1).unwrap())?;
                    args.push(Argument {
                        value: val,
                        name: None,
                        by_ref: false,
                        spread: true,
                    });
                }
            } else if first_text == "*" {
                // Splat
                if children.len() > 1 {
                    let val = walk_expression(children.into_iter().nth(1).unwrap())?;
                    args.push(Argument {
                        value: val,
                        name: None,
                        by_ref: false,
                        spread: true,
                    });
                }
            } else if first_text == "&" || raw_arg.trim_start().starts_with('&') {
                // Block arg
                if raw_arg.trim_start().starts_with('&') {
                    let val = walk_expression(children.into_iter().next().unwrap())?;
                    args.push(Argument::positional(ruby_block_arg_to_lambda(val)));
                } else if children.len() > 1 {
                    let val = walk_expression(children.into_iter().nth(1).unwrap())?;
                    args.push(Argument::positional(ruby_block_arg_to_lambda(val)));
                }
            } else if children.len() >= 2 && children[0].as_rule() == Rule::identifier {
                // Check if keyword arg: identifier ":" expression
                let has_colon = children.iter().any(|c| c.as_str() == ":");
                if has_colon {
                    let name = children[0].as_str().to_string();
                    let val = walk_expression(children.into_iter().last().unwrap())?;
                    args.push(Argument {
                        value: val,
                        name: Some(name),
                        by_ref: false,
                        spread: false,
                    });
                } else {
                    let val = walk_expression(children.into_iter().next().unwrap())?;
                    args.push(Argument::positional(val));
                }
            } else {
                let val = walk_expression(children.into_iter().next().unwrap())?;
                args.push(Argument::positional(val));
            }
        }
    }
    if !pending_hash.is_empty() {
        args.push(Argument::positional(Expression::new(ExprKind::Object(
            pending_hash,
        ))));
    }
    Ok(args)
}

fn ruby_block_arg_to_lambda(expr: Expression) -> Expression {
    if let ExprKind::Call { callee, .. } = &expr.kind {
        if matches!(&callee.kind, ExprKind::Ident(name) if name == "__ruby_method") {
            let param = Param {
                name: "__ruby_proc_arg".to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            };
            let call = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__ruby_proc_call")),
                args: vec![
                    Argument::positional(expr),
                    Argument::positional(Expression::ident("__ruby_proc_arg")),
                ],
                optional: false,
            });
            return Expression::new(ExprKind::Lambda {
                params: vec![param],
                body: LambdaBody::Block(vec![Statement::new(StmtKind::Return(Some(call)))]),
                is_async: false,
                captures: Vec::new(),
            });
        }
    }
    if let ExprKind::Lit(Literal::Str(method)) = &expr.kind {
        let param = Param {
            name: "__ruby_proc_arg".to_string(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        };
        let call = Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("__ruby_proc_arg")),
                field: method.clone(),
                null_safe: false,
            })),
            args: Vec::new(),
            optional: false,
        });
        return Expression::new(ExprKind::Lambda {
            params: vec![param],
            body: LambdaBody::Block(vec![Statement::new(StmtKind::Return(Some(call)))]),
            is_async: false,
            captures: Vec::new(),
        });
    }
    expr
}

// ── Block literal ───────────────────────────────────────────────────────────

fn walk_block_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut params = Vec::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::do_block | Rule::brace_block => {
                for bp in p.into_inner() {
                    match bp.as_rule() {
                        Rule::block_params => {
                            params = walk_block_params(bp)?;
                        }
                        Rule::body => {
                            body = walk_body(bp)?;
                        }
                        _ => {
                            // Statements directly in brace_block
                            let stmt = walk_statement(bp)?;
                            if !matches!(stmt.kind, StmtKind::Empty) {
                                body.push(stmt);
                            }
                        }
                    }
                }
            }
            _ => {}
        }
    }

    apply_implicit_return(&mut body);

    Ok(Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    }))
}

/// Ruby implicit return: last expression in a body becomes a Return.
fn apply_implicit_return(body: &mut Vec<Statement>) {
    if let Some(last) = body.last_mut() {
        if matches!(&last.kind, StmtKind::Expr(_)) {
            if let StmtKind::Expr(e) = std::mem::replace(&mut last.kind, StmtKind::Empty) {
                last.kind = StmtKind::Return(Some(e));
            }
        } else if let StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } = &mut last.kind
        {
            apply_implicit_return(then_body);
            for (_, body) in elifs {
                apply_implicit_return(body);
            }
            if let Some(body) = else_body {
                apply_implicit_return(body);
            }
        }
    }
}

fn lambda_contains_spaceship(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(e) => expr_contains_spaceship(e),
            LambdaBody::Block(stmts) => stmts.iter().any(stmt_contains_spaceship),
        },
        _ => false,
    }
}

fn stmt_contains_spaceship(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::Expr(e) => expr_contains_spaceship(e),
        StmtKind::Return(Some(e)) => expr_contains_spaceship(e),
        StmtKind::Block(stmts) => stmts.iter().any(stmt_contains_spaceship),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            expr_contains_spaceship(cond)
                || then_body.iter().any(stmt_contains_spaceship)
                || elifs.iter().any(|(cond, body)| {
                    expr_contains_spaceship(cond) || body.iter().any(stmt_contains_spaceship)
                })
                || else_body
                    .as_ref()
                    .map(|body| body.iter().any(stmt_contains_spaceship))
                    .unwrap_or(false)
        }
        _ => false,
    }
}

fn expr_contains_spaceship(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Binary { op, left, right } => {
            *op == BinOp::Spaceship || expr_contains_spaceship(left) || expr_contains_spaceship(right)
        }
        ExprKind::Unary { expr, .. } => expr_contains_spaceship(expr),
        ExprKind::Ternary { cond, then, else_ } => {
            expr_contains_spaceship(cond) || expr_contains_spaceship(then) || expr_contains_spaceship(else_)
        }
        ExprKind::Call { callee, args, .. } => {
            expr_contains_spaceship(callee) || args.iter().any(|arg| expr_contains_spaceship(&arg.value))
        }
        ExprKind::Member { object, .. } => expr_contains_spaceship(object),
        ExprKind::Index { object, index, .. } => {
            expr_contains_spaceship(object) || expr_contains_spaceship(index)
        }
        ExprKind::Assign { target, value } => {
            expr_contains_spaceship(target) || expr_contains_spaceship(value)
        }
        ExprKind::Array(elements) => elements.iter().any(|element| {
            element
                .key
                .as_ref()
                .map(expr_contains_spaceship)
                .unwrap_or(false)
                || expr_contains_spaceship(&element.value)
        }),
        ExprKind::Interpolation(parts) => parts.iter().any(|part| match part {
            InterpolPart::Expr(e) | InterpolPart::Formatted(e, _) => expr_contains_spaceship(e),
            InterpolPart::Text(_) => false,
        }),
        ExprKind::Range { start, end, .. } => expr_contains_spaceship(start) || expr_contains_spaceship(end),
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(e) => expr_contains_spaceship(e),
            LambdaBody::Block(stmts) => stmts.iter().any(stmt_contains_spaceship),
        },
        _ => false,
    }
}

fn walk_block_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::block_param_list {
            for bp in p.into_inner() {
                if bp.as_rule() == Rule::block_param_item {
                    let inner = bp.into_inner().next();
                    if let Some(item) = inner {
                        match item.as_rule() {
                            Rule::splat_param => {
                                let name = item
                                    .into_inner()
                                    .find(|c| c.as_rule() == Rule::identifier)
                                    .map(|c| c.as_str().to_string())
                                    .unwrap_or_default();
                                params.push(Param {
                                    name,
                                    type_hint: None,
                                    default: None,
                                    pass_by: PassBy::Value,
                                    is_rest: true,
                                    is_kwargs: false,
                                    is_optional: false,
                                    is_nullable: false,
                                });
                            }
                            Rule::optional_param => {
                                let mut inner = item.into_inner();
                                let name = inner
                                    .next()
                                    .map(|c| c.as_str().to_string())
                                    .unwrap_or_default();
                                let default = inner.find(|c| is_expression_rule(c.as_rule()))
                                    .map(walk_expression)
                                    .transpose()?;
                                params.push(Param {
                                    name,
                                    type_hint: None,
                                    default,
                                    pass_by: PassBy::Value,
                                    is_rest: false,
                                    is_kwargs: false,
                                    is_optional: true,
                                    is_nullable: false,
                                });
                            }
                            Rule::identifier => {
                                params.push(Param {
                                    name: item.as_str().to_string(),
                                    type_hint: None,
                                    default: None,
                                    pass_by: PassBy::Value,
                                    is_rest: false,
                                    is_kwargs: false,
                                    is_optional: false,
                                    is_nullable: false,
                                });
                            }
                            _ => {}
                        }
                    }
                }
            }
        }
    }
    Ok(params)
}

// ── Primary ─────────────────────────────────────────────────────────────────

fn walk_primary(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let source = pair.as_str().trim().to_string();
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }
    if inner.is_empty() {
        if source.starts_with('[') && source.ends_with(']') {
            return Ok(ExprKind::Array(Vec::new()));
        }
        return Ok(ExprKind::Lit(Literal::Null));
    }

    let first = &inner[0];
    match first.as_rule() {
        Rule::array_inner => {
            // Array literal [...]
            walk_array_inner(inner.remove(0))
        }
        Rule::hash_inner => {
            // Hash literal {...}
            walk_hash_inner(inner.remove(0))
        }
        _ => walk_expr_kind(inner.remove(0)),
    }
}

fn walk_array_inner(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let elements = pair
        .into_inner()
        .filter(|p| is_expression_rule(p.as_rule()))
        .map(|p| -> Result<ArrayElement, String> {
            let val = walk_expression(p)?;
            Ok(ArrayElement {
                key: None,
                value: val,
                spread: false,
                by_ref: false,
            })
        })
        .collect::<Result<Vec<_>, _>>()?;
    Ok(ExprKind::Array(elements))
}

fn walk_hash_inner(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut props = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::hash_pair {
            let children: Vec<Pair<Rule>> = p.into_inner().collect();
            if children.len() >= 2 {
                // Could be hash rocket (key => val) or symbol shorthand (key: val)
                let first = &children[0];
                if first.as_rule() == Rule::identifier && children.len() == 2 {
                    // Symbol shorthand: key: val
                    let key =
                        Expression::new(ExprKind::Lit(Literal::Str(first.as_str().to_string())));
                    let val = walk_expression(children.into_iter().nth(1).unwrap())?;
                    props.push(ObjectProperty::KeyValue { key, value: val });
                } else {
                    let key = walk_expression(children[0].clone())?;
                    let val = walk_expression(children.into_iter().last().unwrap())?;
                    props.push(ObjectProperty::KeyValue { key, value: val });
                }
            } else if children.len() == 1 {
                // **expr (double splat)
                let val = walk_expression(children.into_iter().next().unwrap())?;
                props.push(ObjectProperty::Spread(val));
            }
        }
    }
    Ok(ExprKind::Object(props))
}

// ── Interpolated string ─────────────────────────────────────────────────────

fn walk_interpolated_string(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut parts = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::interp_start | Rule::interp_end => {}
            Rule::interp_text => {
                let text = p
                    .as_str()
                    .replace("\\n", "\n")
                    .replace("\\t", "\t")
                    .replace("\\r", "\r")
                    .replace("\\\\", "\\")
                    .replace("\\\"", "\"");
                parts.push(InterpolPart::Text(text));
            }
            Rule::interp_escape => {
                let s = p.as_str();
                let ch = if s.len() >= 2 {
                    match s.chars().nth(1) {
                        Some('n') => "\n",
                        Some('t') => "\t",
                        Some('r') => "\r",
                        Some('\\') => "\\",
                        Some('"') => "\"",
                        Some('#') => "#",
                        _ => s,
                    }
                } else {
                    s
                };
                parts.push(InterpolPart::Text(ch.to_string()));
            }
            Rule::interp_expr => {
                for ip in p.into_inner() {
                    if is_expression_rule(ip.as_rule()) {
                        parts.push(InterpolPart::Expr(walk_expression(ip)?));
                    }
                }
            }
            _ => {}
        }
    }

    // Optimize: if only text parts, concat into single string
    if parts.iter().all(|p| matches!(p, InterpolPart::Text(_))) {
        let s: String = parts
            .iter()
            .map(|p| match p {
                InterpolPart::Text(t) => t.as_str(),
                _ => "",
            })
            .collect();
        return Ok(ExprKind::Lit(Literal::Str(s)));
    }

    Ok(ExprKind::Interpolation(parts))
}

// ── Lambda ──────────────────────────────────────────────────────────────────

fn walk_lambda(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let source = pair.as_str().to_string();
    let mut params = Vec::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::param_list => params = walk_param_list(p)?,
            Rule::body => body = walk_body(p)?,
            _ => {
                // Statements in lambda brace body
                let stmt = walk_statement(p)?;
                if !matches!(stmt.kind, StmtKind::Empty) {
                    body.push(stmt);
                }
            }
        }
    }

    apply_implicit_return(&mut body);
    if let (Some(open), Some(close)) = (source.find('{'), source.rfind('}')) {
        let inner = source[open + 1..close].trim();
        if !inner.is_empty() && !inner.contains(';') && !inner.contains('\n') {
            if let Ok(mut parsed) = RubyParser::parse(Rule::expression, inner) {
                if let Some(expr_pair) = parsed.next() {
                    body = vec![Statement::new(StmtKind::Return(Some(walk_expression(expr_pair)?)))];
                }
            }
        }
    }

    let lambda = Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    });
    Ok(ruby_proc_expr("__ruby_lambda", lambda).kind)
}

fn walk_proc(pair: Pair<Rule>) -> Result<ExprKind, String> {
    for p in pair.into_inner() {
        if p.as_rule() == Rule::block_literal {
            let lambda = walk_block_literal(p)?;
            return Ok(ruby_proc_expr("__ruby_proc", lambda).kind);
        }
    }
    let lambda = Expression::new(ExprKind::Lambda {
        params: Vec::new(),
        body: LambdaBody::Block(Vec::new()),
        is_async: false,
        captures: Vec::new(),
    });
    Ok(ruby_proc_expr("__ruby_proc", lambda).kind)
}

// ── Yield ───────────────────────────────────────────────────────────────────

fn walk_yield(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut args = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::expression_list {
            for ep in p.into_inner() {
                if is_expression_rule(ep.as_rule()) {
                    args.push(walk_expression(ep)?);
                }
            }
        } else if is_expression_rule(p.as_rule()) {
            args.push(walk_expression(p)?);
        }
    }
    // Ruby yield calls the block; emit as Yield for now
    if args.is_empty() {
        Ok(ExprKind::Yield(None))
    } else if args.len() == 1 {
        Ok(ExprKind::Yield(Some(Box::new(
            args.into_iter().next().unwrap(),
        ))))
    } else {
        Ok(ExprKind::Yield(Some(Box::new(Expression::new(
            ExprKind::Array(
                args.into_iter()
                    .map(|a| ArrayElement {
                        key: None,
                        value: a,
                        spread: false,
                        by_ref: false,
                    })
                    .collect(),
            ),
        )))))
    }
}

// ── Defined? ────────────────────────────────────────────────────────────────

fn walk_defined(pair: Pair<Rule>) -> Result<ExprKind, String> {
    // defined?(expr) → check if expr is defined, simplify to !nil
    for p in pair.into_inner() {
        if is_expression_rule(p.as_rule()) {
            let expr = walk_expression(p)?;
            return Ok(ExprKind::Binary {
                op: BinOp::NotEq,
                left: Box::new(expr),
                right: Box::new(Expression::null()),
            });
        }
    }
    Ok(ExprKind::Lit(Literal::Bool(false)))
}

// ── Super ───────────────────────────────────────────────────────────────────

fn walk_super(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut args = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::call_args {
            args = walk_call_args(p)?;
        }
    }
    Ok(ExprKind::SuperCall { method: None, args })
}

// ── If/Unless as expression ─────────────────────────────────────────────────

fn walk_if_expr(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let kind = walk_if(pair)?;
    // Wrap as a ternary-like expression
    if let StmtKind::If {
        cond,
        then_body,
        else_body,
        ..
    } = kind
    {
        let then_val = body_to_expr(then_body);
        let else_val = else_body.map(body_to_expr).unwrap_or(Expression::null());
        Ok(ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(then_val),
            else_: Box::new(else_val),
        })
    } else {
        Ok(ExprKind::Lit(Literal::Null))
    }
}

fn walk_unless_expr(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let kind = walk_unless(pair)?;
    if let StmtKind::If {
        cond,
        then_body,
        else_body,
        ..
    } = kind
    {
        let then_val = body_to_expr(then_body);
        let else_val = else_body.map(body_to_expr).unwrap_or(Expression::null());
        Ok(ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(then_val),
            else_: Box::new(else_val),
        })
    } else {
        Ok(ExprKind::Lit(Literal::Null))
    }
}

fn walk_begin_expr(pair: Pair<Rule>) -> Result<ExprKind, String> {
    // begin..rescue..end as expression — just walk the body
    let kind = walk_begin(pair)?;
    if let StmtKind::Try { body, .. } = kind {
        Ok(body_to_expr(body).kind)
    } else {
        Ok(ExprKind::Lit(Literal::Null))
    }
}

/// Convert a body (list of stmts) to a single expression (last statement value).
fn body_to_expr(mut stmts: Vec<Statement>) -> Expression {
    if stmts.is_empty() {
        return Expression::null();
    }
    let last = stmts.pop().unwrap();
    match last.kind {
        StmtKind::Expr(e) => e,
        StmtKind::Return(Some(e)) => e,
        _ => Expression::null(),
    }
}

// ── Expression list ─────────────────────────────────────────────────────────

fn walk_expr_list_kind(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let inner: Vec<Pair<Rule>> = pair
        .into_inner()
        .filter(|p| is_expression_rule(p.as_rule()))
        .collect();
    if inner.len() == 1 {
        walk_expr_kind(inner.into_iter().next().unwrap())
    } else if inner.is_empty() {
        Ok(ExprKind::Lit(Literal::Null))
    } else {
        let exprs = inner
            .into_iter()
            .map(walk_expression)
            .collect::<Result<Vec<_>, _>>()?;
        Ok(ExprKind::Array(
            exprs
                .into_iter()
                .map(|e| ArrayElement {
                    key: None,
                    value: e,
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        ))
    }
}

fn walk_expr_list_single(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner: Vec<Pair<Rule>> = pair
        .into_inner()
        .filter(|p| is_expression_rule(p.as_rule()))
        .collect();
    if inner.len() == 1 {
        walk_expression(inner.remove(0))
    } else if inner.is_empty() {
        Ok(Expression::null())
    } else {
        let exprs = inner
            .into_iter()
            .map(walk_expression)
            .collect::<Result<Vec<_>, _>>()?;
        Ok(Expression::new(ExprKind::Array(
            exprs
                .into_iter()
                .map(|e| ArrayElement {
                    key: None,
                    value: e,
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        )))
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Helpers
// ════════════════════════════════════════════════════════════════════════════

fn to_span(pair: &Pair<Rule>) -> Span {
    let s = pair.as_span();
    let (sl, sc) = s.start_pos().line_col();
    let (el, ec) = s.end_pos().line_col();
    Span {
        start_line: sl as u32,
        start_col: sc as u32,
        end_line: el as u32,
        end_col: ec as u32,
    }
}

fn negate(expr: Expression) -> Expression {
    Expression::new(ExprKind::Unary {
        op: UnaryOp::Not,
        expr: Box::new(expr),
    })
}

fn next_meaningful<'a>(
    iter: &mut impl Iterator<Item = Pair<'a, Rule>>,
) -> Result<Pair<'a, Rule>, String> {
    for p in iter {
        match p.as_rule() {
            Rule::NEWLINE | Rule::then_kw | Rule::do_kw | Rule::in_kw => continue,
            _ => return Ok(p),
        }
    }
    Err("No more meaningful pairs".into())
}

fn next_rule<'a>(
    iter: &mut impl Iterator<Item = Pair<'a, Rule>>,
    rule: Rule,
) -> Result<Pair<'a, Rule>, String> {
    for p in iter {
        if p.as_rule() == rule {
            return Ok(p);
        }
    }
    Err(format!("Expected {:?}", rule))
}

fn find_rule<'a>(
    iter: impl Iterator<Item = Pair<'a, Rule>>,
    rule: Rule,
) -> Result<Pair<'a, Rule>, String> {
    for p in iter {
        if p.as_rule() == rule {
            return Ok(p);
        }
    }
    Err(format!("Expected {:?}", rule))
}

fn find_rule_from_iter<'a>(
    iter: &mut impl Iterator<Item = Pair<'a, Rule>>,
    rule: Rule,
) -> Result<Pair<'a, Rule>, String> {
    for p in iter {
        if p.as_rule() == rule {
            return Ok(p);
        }
    }
    Err(format!("Expected {:?}", rule))
}

fn is_expression_rule(rule: Rule) -> bool {
    matches!(
        rule,
        Rule::expression
            | Rule::expression_list
            | Rule::ternary_expr
            | Rule::or_expr
            | Rule::and_expr
            | Rule::not_expr
            | Rule::comparison
            | Rule::bitor_expr
            | Rule::bitxor_expr
            | Rule::bitand_expr
            | Rule::shift_expr
            | Rule::range_expr
            | Rule::additive
            | Rule::multiplicative
            | Rule::unary
            | Rule::postfix
            | Rule::primary
            | Rule::integer_literal
            | Rule::float_literal
            | Rule::string_literal
            | Rule::interpolated_string
            | Rule::heredoc
            | Rule::symbol
            | Rule::regex_literal
            | Rule::percent_literal
            | Rule::true_kw
            | Rule::false_kw
            | Rule::nil_kw
            | Rule::self_kw
            | Rule::identifier
            | Rule::constant
            | Rule::constant_path
            | Rule::instance_var
            | Rule::class_var
            | Rule::global_var
            | Rule::yield_expr
            | Rule::defined_expr
            | Rule::super_expr
            | Rule::block_given_expr
            | Rule::lambda_literal
            | Rule::proc_literal
            | Rule::if_expr
            | Rule::unless_expr
            | Rule::begin_expr
    )
}

fn is_op_rule(rule: Rule) -> bool {
    matches!(
        rule,
        Rule::additive_op
            | Rule::multiplicative_op
            | Rule::shift_op
            | Rule::comparison_op
            | Rule::range_op
            | Rule::aug_assign_op
    )
}

fn parse_comparison_op(s: &str) -> BinOp {
    match s {
        "==" => BinOp::Eq,
        "!=" => BinOp::NotEq,
        "<" => BinOp::Lt,
        ">" => BinOp::Gt,
        "<=" => BinOp::LtEq,
        ">=" => BinOp::GtEq,
        "<=>" => BinOp::Spaceship,
        "===" => BinOp::StrictEq,
        _ => BinOp::Eq,
    }
}

fn parse_binop(s: &str) -> BinOp {
    match s {
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "%" => BinOp::Mod,
        "**" => BinOp::Pow,
        "<<" => BinOp::Shl,
        ">>" => BinOp::Shr,
        "|" => BinOp::BitOr,
        "^" => BinOp::BitXor,
        "&" => BinOp::BitAnd,
        _ => BinOp::Add,
    }
}

fn ruby_int_expr(value: i64) -> Expression {
    Expression::new(ExprKind::Lit(Literal::Int(value)))
}

fn ruby_call_expr(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn ruby_proc_expr(name: &str, lambda: Expression) -> Expression {
    let (arity, has_rest) = match &lambda.kind {
        ExprKind::Lambda { params, .. } => (
            params.iter().filter(|p| !p.is_rest).count() as i64,
            params.iter().any(|p| p.is_rest),
        ),
        _ => (0, false),
    };
    let param_count = match &lambda.kind {
        ExprKind::Lambda { params, .. } => params.len() as i64,
        _ => arity,
    };
    ruby_call_expr(
        name,
        vec![
            lambda,
            Expression::new(ExprKind::Lit(Literal::Int(arity))),
            Expression::new(ExprKind::Lit(Literal::Bool(has_rest))),
            Expression::new(ExprKind::Lit(Literal::Int(param_count))),
        ],
    )
}

fn ruby_add_expr(left: Expression, right: Expression) -> Expression {
    ruby_call_expr("__ruby_op_add", vec![left, right])
}

fn ruby_sub_expr(left: Expression, right: Expression) -> Expression {
    ruby_call_expr("__ruby_op_sub", vec![left, right])
}

fn is_negative_one_expr(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr
        } if matches!(&expr.kind, ExprKind::Lit(Literal::Int(1)))
    )
}

fn is_negative_int_expr(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr
        } if matches!(&expr.kind, ExprKind::Lit(Literal::Int(_)))
    )
}

fn literal_int_value(expr: &Expression) -> Option<i64> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(v)) => Some(*v),
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        } => {
            if let ExprKind::Lit(Literal::Int(v)) = &expr.kind {
                Some(-*v)
            } else {
                None
            }
        }
        _ => None,
    }
}

fn ruby_slice_returns_nil(receiver: &Expression, method_name: &str, args: &[Argument]) -> bool {
    if method_name != "slice" || args.len() != 2 {
        return false;
    }
    if is_negative_int_expr(&args[1].value) {
        return true;
    }
    if let ExprKind::Array(elements) = &receiver.kind {
        if let Some(start) = literal_int_value(&args[0].value) {
            return start >= elements.len() as i64;
        }
    }
    false
}

fn ruby_range_exclusive_end(end: Expression, inclusive: bool) -> Expression {
    if !inclusive {
        return end;
    }
    if is_negative_one_expr(&end) {
        ruby_int_expr(i32::MAX as i64)
    } else {
        ruby_add_expr(end, ruby_int_expr(1))
    }
}

fn normalize_ruby_slice_call(method_name: &mut String, args: &mut Vec<Argument>) {
    if method_name != "slice" && method_name != "slice!" {
        return;
    }

    if args.len() == 1 {
        if let ExprKind::Range {
            start,
            end,
            inclusive,
        } = args[0].value.clone().kind
        {
            let start = *start;
            let exclusive_end = ruby_range_exclusive_end(*end, inclusive);
            if method_name == "slice!" {
                let count = ruby_sub_expr(exclusive_end, start.clone());
                args.clear();
                args.push(Argument::positional(start));
                args.push(Argument::positional(count));
            } else {
                args.clear();
                args.push(Argument::positional(start));
                args.push(Argument::positional(exclusive_end));
            }
        }
    } else if args.len() == 2 && method_name == "slice" {
        let start = args[0].value.clone();
        let len = args[1].value.clone();
        args[1].value = ruby_add_expr(start, len);
    }
}

fn parse_ruby_int(s: &str) -> Result<ExprKind, String> {
    let s = s.replace('_', "");
    if s.starts_with("0x") || s.starts_with("0X") {
        Ok(ExprKind::Lit(Literal::Int(
            i64::from_str_radix(&s[2..], 16).unwrap_or(0),
        )))
    } else if s.starts_with("0o") || s.starts_with("0O") {
        Ok(ExprKind::Lit(Literal::Int(
            i64::from_str_radix(&s[2..], 8).unwrap_or(0),
        )))
    } else if s.starts_with("0b") || s.starts_with("0B") {
        Ok(ExprKind::Lit(Literal::Int(
            i64::from_str_radix(&s[2..], 2).unwrap_or(0),
        )))
    } else {
        Ok(ExprKind::Lit(Literal::Int(s.parse().unwrap_or(0))))
    }
}

fn parse_ruby_float(s: &str) -> Result<ExprKind, String> {
    let s = s.replace('_', "");
    Ok(ExprKind::Lit(Literal::Float(s.parse().unwrap_or(0.0))))
}

fn parse_ruby_string(s: &str) -> String {
    let s = if s.starts_with("'''") {
        &s[3..s.len() - 3]
    } else if s.starts_with('\'') {
        &s[1..s.len() - 1]
    } else {
        s
    };
    s.replace("\\n", "\n")
        .replace("\\r", "\r")
        .replace("\\t", "\t")
        .replace("\\'", "'")
        .replace("\\\\", "\\")
}

fn parse_heredoc(s: &str) -> String {
    // <<~TAG\ncontent\nTAG  or  <<TAG\ncontent\nTAG
    let squiggly = s.starts_with("<<~");
    let prefix_len = if squiggly { 3 } else { 2 };
    let rest = &s[prefix_len..];
    // Find the tag name (up to newline)
    if let Some(nl) = rest.find('\n') {
        let tag = rest[..nl].trim();
        let content = &rest[nl + 1..];
        // Strip trailing TAG line
        let body = if let Some(pos) = content.rfind(tag) {
            &content[..pos]
        } else {
            content
        };
        if squiggly {
            // Strip common leading whitespace
            let lines: Vec<&str> = body.lines().collect();
            let min_indent = lines
                .iter()
                .filter(|l| !l.trim().is_empty())
                .map(|l| l.len() - l.trim_start().len())
                .min()
                .unwrap_or(0);
            lines
                .iter()
                .map(|l| {
                    if l.len() > min_indent {
                        &l[min_indent..]
                    } else {
                        l.trim()
                    }
                })
                .collect::<Vec<_>>()
                .join("\n")
        } else {
            body.to_string()
        }
    } else {
        s.to_string()
    }
}

fn walk_percent_literal(s: &str) -> ExprKind {
    // %w[a b c] → array of strings
    // %W[a #{x} c] → array of interpolated strings
    // %i[a b c] → array of symbols (strings)
    // %I[a #{x} c] → array of interpolated symbols (strings)
    // %q[...] → single-quoted string
    // %Q[...] or %[...] → double-quoted string
    let (kind, interpolate, rest) = if s.starts_with("%w") || s.starts_with("%i") {
        ("array", false, &s[2..])
    } else if s.starts_with("%W") || s.starts_with("%I") {
        ("array", true, &s[2..])
    } else if s.starts_with("%q") || s.starts_with("%Q") {
        ("string", s.starts_with("%Q"), &s[2..])
    } else {
        ("string", true, &s[1..])
    };

    // Strip delimiters
    let body = if rest.len() >= 2 {
        &rest[1..rest.len() - 1]
    } else {
        rest
    };

    if kind == "array" {
        let words: Vec<ArrayElement> = ruby_percent_words(body, interpolate)
            .into_iter()
            .map(|w| ArrayElement {
                key: None,
                value: if interpolate && w.starts_with("#{") && w.ends_with('}') {
                    Expression::ident(&w[2..w.len() - 1])
                } else if !interpolate && w.starts_with("#{") && w.ends_with('}') {
                    Expression::new(ExprKind::Lit(Literal::Str(format!("\\{}", w))))
                } else {
                    Expression::new(ExprKind::Lit(Literal::Str(w)))
                },
                spread: false,
                by_ref: false,
            })
            .collect();
        ExprKind::Array(words)
    } else {
        ExprKind::Lit(Literal::Str(body.to_string()))
    }
}

fn ruby_percent_words(body: &str, interpolate: bool) -> Vec<String> {
    let mut words = Vec::new();
    let mut cur = String::new();
    let mut chars = body.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch.is_whitespace() {
            if !cur.is_empty() {
                words.push(std::mem::take(&mut cur));
            }
            continue;
        }
        if ch == '\\' {
            match chars.next() {
                Some(' ') => cur.push(' '),
                Some('n') if interpolate => cur.push('\n'),
                Some(other) => cur.push(other),
                None => cur.push('\\'),
            }
        } else {
            cur.push(ch);
        }
    }
    if !cur.is_empty() {
        words.push(cur);
    }
    words
}
