/// A very small, line-oriented AST for Python used by the early compiler.
#[derive(Debug, Clone)]
pub enum Expr {
    Int(i32),
    Float(f64),
    Str(String),
    Bool(bool),
    None,
    Ident(String),
    Unary { op: String, expr: Box<Expr> },
    Binary { op: String, left: Box<Expr>, right: Box<Expr> },
    List(Vec<Expr>),
    Tuple(Vec<Expr>),
    Dict(Vec<(Expr, Expr)>),
    Index { obj: Box<Expr>, idx: Box<Expr> },
    Call { func: Box<Expr>, args: Vec<Expr> },
}

#[derive(Debug, Clone)]
pub enum Stmt {
    Assign { name: String, expr: Expr },
    Expr { expr: Expr },
    Print { args: Vec<Expr> },
    Return { expr: Expr },
    Break,
    Continue,
        For { target: String, iter: Expr, body: Vec<Stmt> },
    If { cond: Expr, then_branch: Vec<Stmt>, else_branch: Option<Vec<Stmt>> },
    While { cond: Expr, body: Vec<Stmt> },
    /// Multi-line or single-line function def
    Function { name: String, args: Vec<String>, body: Vec<Stmt> },
}

#[derive(Debug, Clone)]
pub struct Program {
    pub stmts: Vec<Stmt>,
}

/// Parse source into a line-oriented AST. This intentionally keeps expressions
/// as raw strings so the compiler can reuse the existing quick expression
/// handling while we migrate to a richer expression AST later.
pub fn parse(source: &str) -> Result<Program, String> {
    // Split into (indent, text) pairs
    let mut lines: Vec<(usize, String)> = Vec::new();
    for raw in source.lines() {
        let mut indent = 0usize;
        for c in raw.chars() {
            if c == ' ' { indent += 1; } else { break; }
        }
        let text = raw[indent..].trim_end().to_string();
        lines.push((indent, text));
    }

    // Recursive parser that consumes a contiguous region of lines with minimal indent >= base_indent
    fn parse_block(lines: &[(usize, String)], start: usize, base_indent: usize) -> Result<(Vec<Stmt>, usize), String> {
        // helper: normalize any `Stmt::Expr { Expr::Ident("break"/"continue") }` into Stmt::Break/Continue
        fn normalize_stmts(stmts: &mut Vec<Stmt>) {
            for s in stmts.iter_mut() {
                match s {
                    Stmt::Expr { expr } => {
                        // no-op here; kept for symmetry with earlier code
                        let _ = expr;
                    }
                    _ => {}
                }
            }
        }

        let mut i = start;
        let mut stmts: Vec<Stmt> = Vec::new();
        while i < lines.len() {
            let (indent, ref text) = lines[i];
            if indent < base_indent { break; }
            let line = text.trim();
            if line.is_empty() { i += 1; continue; }
            if line.starts_with('#') { i += 1; continue; }

            // Single-line if: `if <cond>: <stmt>`
            if line.starts_with("if ") {
                if let Some(colon_pos) = line.find(':') {
                    if colon_pos != line.len() - 1 {
                        let cond_text = line[3..colon_pos].trim();
                        let cond = parse_expr(cond_text)?;
                        let body_text = line[colon_pos+1..].trim();
                        let mut then_branch: Vec<Stmt> = Vec::new();
                        if body_text.starts_with("return ") {
                            let expr_text = body_text[7..].trim();
                            then_branch.push(Stmt::Return { expr: parse_expr(expr_text)? });
                        } else if body_text == "break" {
                            then_branch.push(Stmt::Break);
                        } else if body_text == "continue" {
                            then_branch.push(Stmt::Continue);
                        } else if body_text.starts_with("print(") && body_text.ends_with(")") {
                            let inner = &body_text[6..body_text.len()-1];
                            let mut args_text: Vec<String> = Vec::new();
                            let mut start_idx = 0usize; let mut depth: i32 = 0; let mut in_string: Option<char> = None; let mut esc = false;
                            for (j, ch) in inner.char_indices() {
                                if esc { esc = false; continue; }
                                match ch {
                                    '\\' => esc = true,
                                    '"' | '\'' => { if let Some(q) = in_string { if q == ch { in_string = None; } } else { in_string = Some(ch); } }
                                    '[' | '(' if in_string.is_none() => depth += 1,
                                    ']' | ')' if in_string.is_none() => depth -= 1,
                                    ',' if depth == 0 && in_string.is_none() => { let token = inner[start_idx..j].trim(); if !token.is_empty() { args_text.push(token.to_string()); } start_idx = j + 1; }
                                    _ => {}
                                }
                            }
                            if start_idx <= inner.len() { let token = inner[start_idx..].trim(); if !token.is_empty() { args_text.push(token.to_string()); } }
                            let mut args: Vec<Expr> = Vec::new(); for a in args_text.iter() { match parse_expr(a) { Ok(v) => args.push(v), Err(e) => { eprintln!("vybe_parser_python: parse_expr failed for print-arg {:?} in line {:?}: {}", a, body_text, e); return Err(e); } } }
                            then_branch.push(Stmt::Print { args });
                        } else if let Some(eq) = body_text.find('=') {
                            let name = body_text[..eq].trim().to_string();
                            let expr_text = body_text[eq+1..].trim();
                            let expr = parse_expr(expr_text)?;
                            then_branch.push(Stmt::Assign { name, expr });
                        } else {
                            then_branch.push(Stmt::Expr { expr: parse_expr(body_text)? });
                        }
                        stmts.push(Stmt::If { cond, then_branch, else_branch: None });
                        i += 1; continue;
                    }
                }
            }

            // Block-starter
            if line.ends_with(":") {
                let header = line.trim_end_matches(':').trim();
                if header.starts_with("def ") {
                    if let Some(open) = header.find('(') {
                        if let Some(close) = header.find(')') {
                            let name = header[4..open].trim().to_string();
                            let args_text = &header[open+1..close];
                            let args: Vec<String> = args_text.split(',').map(|s| s.trim().to_string()).filter(|s| !s.is_empty()).collect();
                            let (mut body, consumed) = parse_block(lines, i+1, indent + 1)?;
                            normalize_stmts(&mut body);
                            stmts.push(Stmt::Function { name, args, body });
                            i = consumed; continue;
                        }
                    }
                }
                if header.starts_with("if ") {
                    let cond_text = header[3..].trim();
                    let cond = parse_expr(cond_text)?;
                    let (then_body, consumed) = parse_block(lines, i+1, indent + 1)?;
                    i = consumed;
                    // Check for an `else:` block immediately following at the same indent
                    let mut else_branch: Option<Vec<Stmt>> = None;
                    if i < lines.len() {
                        let (next_indent, ref next_text) = lines[i];
                        if next_indent == indent {
                            let nt = next_text.trim();
                            if nt.starts_with("else") && nt.ends_with(":" ) {
                                let (ebody, consumed2) = parse_block(lines, i+1, indent + 1)?;
                                else_branch = Some(ebody);
                                i = consumed2;
                            }
                        }
                    }
                    stmts.push(Stmt::If { cond, then_branch: then_body, else_branch });
                    continue;
                }
                if header.starts_with("while ") {
                    let cond_text = header[6..].trim();
                    let cond = parse_expr(cond_text)?;
                    let (mut body, consumed) = parse_block(lines, i+1, indent + 1)?;
                    normalize_stmts(&mut body);
                    stmts.push(Stmt::While { cond, body });
                    i = consumed; continue;
                }
                if header.starts_with("for ") {
                    // expect `for <name> in <expr>`
                    let rest = header[4..].trim();
                    if let Some(in_pos) = rest.find(" in ") {
                        let target = rest[..in_pos].trim().to_string();
                        let iter_text = rest[in_pos+4..].trim();
                        let iter_expr = parse_expr(iter_text)?;
                        let (mut body, consumed) = parse_block(lines, i+1, indent + 1)?;
                        normalize_stmts(&mut body);
                        stmts.push(Stmt::For { target, iter: iter_expr, body });
                        i = consumed; continue;
                    }
                }
                // other block -> treat as expr
                stmts.push(Stmt::Expr { expr: Expr::Str(header.to_string()) });
                i += 1; continue;
            }

            // Assignment
            if let Some(eq) = line.find('=') {
                let name = line[..eq].trim().to_string();
                let expr_text = line[eq+1..].trim();
                let expr = parse_expr(expr_text)?;
                stmts.push(Stmt::Assign { name, expr });
                i += 1; continue;
            }

            // Return
            if line.starts_with("return ") {
                let expr_text = line[7..].trim();
                let expr = parse_expr(expr_text)?;
                stmts.push(Stmt::Return { expr });
                i += 1; continue;
            }

            if line == "break" { stmts.push(Stmt::Break); i += 1; continue; }
            if line == "continue" { stmts.push(Stmt::Continue); i += 1; continue; }

            // print
            if line.starts_with("print(") && line.ends_with(")") {
                let inner = &line[6..line.len()-1];
                let mut args_text: Vec<String> = Vec::new();
                let mut start_idx = 0usize; let mut depth: i32 = 0; let mut in_string: Option<char> = None; let mut esc = false;
                for (j, ch) in inner.char_indices() {
                    if esc { esc = false; continue; }
                    match ch {
                        '\\' => esc = true,
                        '"' | '\'' => { if let Some(q) = in_string { if q == ch { in_string = None; } } else { in_string = Some(ch); } }
                        '[' | '(' if in_string.is_none() => depth += 1,
                        ']' | ')' if in_string.is_none() => depth -= 1,
                        ',' if depth == 0 && in_string.is_none() => { let token = inner[start_idx..j].trim(); if !token.is_empty() { args_text.push(token.to_string()); } start_idx = j + 1; }
                        _ => {}
                    }
                }
                if start_idx <= inner.len() { let token = inner[start_idx..].trim(); if !token.is_empty() { args_text.push(token.to_string()); } }
                let mut args: Vec<Expr> = Vec::new(); for a in args_text.iter() { match parse_expr(a) { Ok(v) => args.push(v), Err(e) => { eprintln!("vybe_parser_python: parse_expr failed for print-arg {:?} in line {:?}: {}", a, line, e); return Err(e); } } }
                stmts.push(Stmt::Print { args }); i += 1; continue;
            }

            // Fallback expression
            let expr = match parse_expr(line) { Ok(v) => v, Err(e) => { eprintln!("vybe_parser_python: parse_expr failed for top-level line {:?}: {}", line, e); return Err(e); } };
            match &expr {
                Expr::Ident(s) if s == "break" => { stmts.push(Stmt::Break); }
                Expr::Ident(s) if s == "continue" => { stmts.push(Stmt::Continue); }
                _ => { stmts.push(Stmt::Expr { expr }); }
            }
            i += 1;
        }
        Ok((stmts, i))
    }

    // parse_expr implementation
    fn parse_expr(s: &str) -> Result<Expr, String> {
        let s = s.trim();
        if s.is_empty() {
            eprintln!("vybe_parser_python: parse_expr received empty string while parsing.\nBacktrace:\n{:?}", std::backtrace::Backtrace::capture());
            return Err(format!("empty expression while parsing: {:?}", s));
        }

        // List literal
        if s.starts_with('[') && s.ends_with(']') {
            let inner = &s[1..s.len()-1];
            let mut elems: Vec<Expr> = Vec::new();
            let mut start = 0usize; let mut depth: i32 = 0; let mut in_string: Option<char> = None; let mut esc=false;
            for (i,ch) in inner.char_indices() {
                if esc { esc=false; continue; }
                match ch {
                    '\\' => esc=true,
                    '"' | '\'' => { if let Some(q)=in_string { if q==ch { in_string=None; } } else { in_string=Some(ch); } }
                    '[' if in_string.is_none() => depth+=1,
                    ']' if in_string.is_none() => depth-=1,
                    ',' if depth==0 && in_string.is_none() => { let token = inner[start..i].trim(); if !token.is_empty() { match parse_expr(token) { Ok(v) => elems.push(v), Err(e) => { eprintln!("vybe_parser_python: parse_expr failed for list token {:?} inside {:?}: {}", token, s, e); return Err(e); } } } start=i+1; }
                    _=>{}
                }
            }
            if start<=inner.len() { let token = inner[start..].trim(); if !token.is_empty() { match parse_expr(token) { Ok(v) => elems.push(v), Err(e) => { eprintln!("vybe_parser_python: parse_expr failed for list final token {:?} inside {:?}: {}", token, s, e); return Err(e); } } } }
            return Ok(Expr::List(elems));
        }

        // Dict literal: {k: v, ...}
        if s.starts_with('{') && s.ends_with('}') {
            let inner = &s[1..s.len()-1];
            let mut items: Vec<(Expr, Expr)> = Vec::new();
            let mut start = 0usize; let mut depth: i32 = 0; let mut in_string: Option<char> = None; let mut esc=false;
            for (i,ch) in inner.char_indices() {
                if esc { esc=false; continue; }
                match ch {
                    '\\' => esc=true,
                    '"' | '\'' => { if let Some(q)=in_string { if q==ch { in_string=None; } } else { in_string=Some(ch); } }
                    '{'|'['|'(' if in_string.is_none() => depth+=1,
                    '}'|']'|')' if in_string.is_none() => depth-=1,
                    ',' if depth==0 && in_string.is_none() => { let token = inner[start..i].trim(); if !token.is_empty() { if let Some(colon) = token.find(':') { let k = token[..colon].trim(); let v = token[colon+1..].trim(); let key = match parse_expr(k) { Ok(x) => x, Err(e) => { eprintln!("vybe_parser_python: parse_expr failed for dict key {:?} inside {:?}: {}", k, s, e); return Err(e); } }; let val = match parse_expr(v) { Ok(x) => x, Err(e) => { eprintln!("vybe_parser_python: parse_expr failed for dict value {:?} inside {:?}: {}", v, s, e); return Err(e); } }; items.push((key, val)); } } start=i+1; }
                    _=>{}
                }
            }
            if start<=inner.len() { let token = inner[start..].trim(); if !token.is_empty() { if let Some(colon) = token.find(':') { let k = token[..colon].trim(); let v = token[colon+1..].trim(); let key = match parse_expr(k) { Ok(x) => x, Err(e) => { eprintln!("vybe_parser_python: parse_expr failed for dict final key {:?} inside {:?}: {}", k, s, e); return Err(e); } }; let val = match parse_expr(v) { Ok(x) => x, Err(e) => { eprintln!("vybe_parser_python: parse_expr failed for dict final value {:?} inside {:?}: {}", v, s, e); return Err(e); } }; items.push((key, val)); } } }
            return Ok(Expr::Dict(items));
        }

        // Tuple literal: (a, b)
        if s.starts_with('(') && s.ends_with(')') {
            let inner = &s[1..s.len()-1];
            // empty tuple
            if inner.trim().is_empty() { return Ok(Expr::Tuple(Vec::new())); }
            let mut elems: Vec<Expr> = Vec::new();
            let mut start = 0usize; let mut depth: i32 = 0; let mut in_string: Option<char> = None; let mut esc=false;
            for (i,ch) in inner.char_indices() {
                if esc { esc=false; continue; }
                match ch {
                    '\\' => esc=true,
                    '"' | '\'' => { if let Some(q)=in_string { if q==ch { in_string=None; } } else { in_string=Some(ch); } }
                    '('|'[' if in_string.is_none() => depth+=1,
                    ')'|']' if in_string.is_none() => depth-=1,
                    ',' if depth==0 && in_string.is_none() => { let token = inner[start..i].trim(); if !token.is_empty() { match parse_expr(token) { Ok(v) => elems.push(v), Err(e) => { eprintln!("vybe_parser_python: parse_expr failed for tuple token {:?} inside {:?}: {}", token, s, e); return Err(e); } } } start=i+1; }
                    _=>{}
                }
            }
            if start<=inner.len() { let token = inner[start..].trim(); if !token.is_empty() { match parse_expr(token) { Ok(v) => elems.push(v), Err(e) => { eprintln!("vybe_parser_python: parse_expr failed for tuple final token {:?} inside {:?}: {}", token, s, e); return Err(e); } } } }
            return Ok(Expr::Tuple(elems));
        }

        // Indexing: support chained forms like `a[b]` and `a[b][c]`.
        if let Some(br) = s.find('[') {
            // find matching close for the first '['
            let mut depth = 0i32; let mut in_string: Option<char> = None; let mut esc=false; let mut close_idx: Option<usize> = None;
            for (i,ch) in s.char_indices() {
                if i < br { continue; }
                if esc { esc=false; continue; }
                match ch {
                    '\\' => esc=true,
                    '"' | '\'' => { if let Some(q)=in_string { if q==ch { in_string=None; } } else { in_string=Some(ch); } }
                    '[' if in_string.is_none() => depth += 1,
                    ']' if in_string.is_none() => { depth -= 1; if depth==0 { close_idx = Some(i); break; } }
                    _=>{}
                }
            }
            if let Some(close) = close_idx {
                // If there are chained indexes, iteratively build nested Index Exprs
                if close < s.len()-1 {
                    // Start by parsing the primary `obj[index]` as a sub-expression
                    let first_chunk = &s[..close+1];
                    let mut expr = match parse_expr(first_chunk) { Ok(v) => v, Err(e) => { eprintln!("vybe_parser_python: parse_expr failed for index first_chunk {:?} inside {:?}: {}", first_chunk, s, e); return Err(e); } };
                    let mut rem = s[close+1..].trim();
                    while rem.starts_with('[') {
                        // find matching close in the remainder
                        let mut d = 0i32; let mut in_s: Option<char> = None; let mut e=false; let mut found: Option<usize> = None;
                        for (j,ch) in rem.char_indices() {
                            if e { e=false; continue; }
                            match ch {
                                '\\' => e=true,
                                '"' | '\'' => { if let Some(q)=in_s { if q==ch { in_s=None; } } else { in_s=Some(ch); } }
                                '[' if in_s.is_none() => d+=1,
                                ']' if in_s.is_none() => { d-=1; if d==0 { found = Some(j); break; } }
                                _=>{}
                            }
                        }
                        if let Some(f) = found {
                            let idx_text = rem[1..f].trim();
                            let idx_expr = match parse_expr(idx_text) { Ok(v) => v, Err(e) => { eprintln!("vybe_parser_python: parse_expr failed for index token {:?} inside {:?}: {}", idx_text, s, e); return Err(e); } };
                            expr = Expr::Index { obj: Box::new(expr), idx: Box::new(idx_expr) };
                            rem = rem[f+1..].trim();
                        } else { break; }
                    }
                    if rem.is_empty() { return Ok(expr); }
                } else {
                    // single index covering the whole string
                    let obj = s[..br].trim(); let idx = s[br+1..close].trim();
                    let o = parse_expr(obj)?; let i = parse_expr(idx)?;
                    return Ok(Expr::Index { obj: Box::new(o), idx: Box::new(i) });
                }
            }
        }

        // Call: foo(a,b) — find matching closing paren for the first '('
        if let Some(open) = s.find('(') {
            let mut depth = 0i32; let mut in_string: Option<char> = None; let mut esc=false; let mut close_idx: Option<usize> = None;
            for (i,ch) in s.char_indices() {
                if i < open { continue; }
                if esc { esc=false; continue; }
                match ch {
                    '\\' => esc=true,
                    '"' | '\'' => { if let Some(q)=in_string { if q==ch { in_string=None; } } else { in_string=Some(ch); } }
                    '(' if in_string.is_none() => depth += 1,
                    ')' if in_string.is_none() => { depth -= 1; if depth==0 { close_idx = Some(i); break; } }
                    _=>{}
                }
            }
                if let Some(close) = close_idx {
                if close == s.len()-1 {
                    let func = s[..open].trim(); let inner = &s[open+1..close];
                    // Special-case: treat `not(<expr>)` and `-(<expr>)` as unary operators,
                    // not as a call to an identifier named "not" or "-". This avoids
                    // parsing `not (x)` as `Call(Ident("not"), [x])` which then becomes
                    // an unknown identifier at compile time. Handle these here and
                    // delegate to the unary handling.
                    if func == "not" {
                        return Ok(Expr::Unary { op: "not".to_string(), expr: Box::new(parse_expr(inner)?) });
                    }
                    if func == "-" {
                        return Ok(Expr::Unary { op: "-".to_string(), expr: Box::new(parse_expr(inner)?) });
                    }
                    let mut args: Vec<Expr> = Vec::new();
                    if !inner.trim().is_empty() {
                        let mut start=0usize; let mut depth=0i32; let mut in_string: Option<char>=None; let mut esc=false;
                        for (i,ch) in inner.char_indices() {
                            if esc { esc=false; continue; }
                            match ch {
                                '\\' => esc=true,
                                '"' | '\'' => { if let Some(q)=in_string { if q==ch { in_string=None; } } else { in_string=Some(ch); } }
                                '('|'[' if in_string.is_none() => depth+=1,
                                ')'|']' if in_string.is_none() => depth-=1,
                                ',' if depth==0 && in_string.is_none() => { let token = inner[start..i].trim(); if !token.is_empty() { match parse_expr(token) { Ok(v) => args.push(v), Err(e) => { eprintln!("vybe_parser_python: parse_expr failed for call-arg token {:?} inside {:?}: {}", token, s, e); return Err(e); } } } start=i+1; }
                                _=>{}
                            }
                        }
                        if start<=inner.len() { let token = inner[start..].trim(); if !token.is_empty() { match parse_expr(token) { Ok(v) => args.push(v), Err(e) => { eprintln!("vybe_parser_python: parse_expr failed for call-arg final token {:?} inside {:?}: {}", token, s, e); return Err(e); } } } }
                    }
                    let fexpr = match parse_expr(func) { Ok(v) => v, Err(e) => { eprintln!("vybe_parser_python: parse_expr failed for call func {:?} inside {:?}: {}", func, s, e); return Err(e); } };
                    return Ok(Expr::Call { func: Box::new(fexpr), args });
                }
            }
        }

        // Unary operators: not, unary -
        if s.starts_with("not ") {
            let inner = s[4..].trim();
            return Ok(Expr::Unary { op: "not".to_string(), expr: Box::new(parse_expr(inner)?) });
        }
        if s.starts_with('-') {
            let inner = s[1..].trim();
            // unary minus
            return Ok(Expr::Unary { op: "-".to_string(), expr: Box::new(parse_expr(inner)?) });
        }

        // Logical OR (lowest precedence)
        {
            let mut depth=0i32; let mut in_string: Option<char>=None; let mut esc=false;
            let mut idx_opt: Option<usize> = None;
            for (i,ch) in s.char_indices() {
                if esc { esc=false; continue; }
                match ch {
                    '\\' => esc=true,
                    '"' | '\'' => { if let Some(q)=in_string { if q==ch { in_string=None; } } else { in_string=Some(ch); } }
                    '('|'[' if in_string.is_none() => depth+=1,
                    ')'|']' if in_string.is_none() => depth-=1,
                    ' ' if in_string.is_none() && depth==0 => {
                        if s[i..].starts_with(" or ") { idx_opt = Some(i); break; }
                    }
                    _=>{}
                }
            }
            if let Some(i) = idx_opt {
                let left = &s[..i]; let right = &s[i+4..];
                return Ok(Expr::Binary { op: "or".to_string(), left: Box::new(parse_expr(left)?), right: Box::new(parse_expr(right)?) });
            }
        }

        // Logical AND
        {
            let mut depth=0i32; let mut in_string: Option<char>=None; let mut esc=false;
            let mut idx_opt: Option<usize> = None;
            for (i,ch) in s.char_indices() {
                if esc { esc=false; continue; }
                match ch {
                    '\\' => esc=true,
                    '"' | '\'' => { if let Some(q)=in_string { if q==ch { in_string=None; } } else { in_string=Some(ch); } }
                    '('|'[' if in_string.is_none() => depth+=1,
                    ')'|']' if in_string.is_none() => depth-=1,
                    ' ' if in_string.is_none() && depth==0 => {
                        if s[i..].starts_with(" and ") { idx_opt = Some(i); break; }
                    }
                    _=>{}
                }
            }
            if let Some(i) = idx_opt {
                let left = &s[..i]; let right = &s[i+5..];
                return Ok(Expr::Binary { op: "and".to_string(), left: Box::new(parse_expr(left)?), right: Box::new(parse_expr(right)?) });
            }
        }

        // Comparisons
        {
            let comps = ["<=", ">=", "==", "!=", "<", ">"];
            for &op in comps.iter() {
                let mut depth=0i32; let mut in_string: Option<char>=None; let mut esc=false;
                let mut found: Option<usize> = None;
                for (i,ch) in s.char_indices() {
                    if esc { esc=false; continue; }
                    match ch {
                        '\\' => esc=true,
                            '"' | '\'' => { if let Some(q)=in_string { if q==ch { in_string=None; } } else { in_string=Some(ch); } }
                        '('|'[' if in_string.is_none() => depth+=1,
                        ')'|']' if in_string.is_none() => depth-=1,
                        _=>{}
                    }
                    if in_string.is_none() && depth==0 {
                        if s[i..].starts_with(op) { found = Some(i); break; }
                    }
                }
                if let Some(i) = found {
                    let left = &s[..i]; let right = &s[i+op.len()..];
                    return Ok(Expr::Binary { op: op.to_string(), left: Box::new(parse_expr(left)?), right: Box::new(parse_expr(right)?) });
                }
            }
        }

        // Binary ops: handle + and - at top level, then * /
        for op in ["+", "-"].iter() {
            let mut depth=0i32; let mut in_string: Option<char>=None; let mut esc=false;
            for (i,ch) in s.char_indices() {
                if esc { esc=false; continue; }
                match ch {
                    '\\' => esc=true,
                    '"' | '\'' => { if let Some(q)=in_string { if q==ch { in_string=None; } } else { in_string=Some(ch); } }
                    '('|'[' if in_string.is_none() => depth+=1,
                    ')'|']' if in_string.is_none() => depth-=1,
                    c if c.to_string()==*op && depth==0 && in_string.is_none() => {
                        let left = &s[..i]; let right = &s[i+op.len()..];
                        return Ok(Expr::Binary { op: op.to_string(), left: Box::new(parse_expr(left)?), right: Box::new(parse_expr(right)?) });
                    }
                    _=>{}
                }
            }
        }
        for op in ["*", "/"].iter() {
            let mut depth=0i32; let mut in_string: Option<char>=None; let mut esc=false;
            for (i,ch) in s.char_indices() {
                if esc { esc=false; continue; }
                match ch {
                    '\\' => esc=true,
                    '"' | '\'' => { if let Some(q)=in_string { if q==ch { in_string=None; } } else { in_string=Some(ch); } }
                    '('|'[' if in_string.is_none() => depth+=1,
                    ')'|']' if in_string.is_none() => depth-=1,
                    c if c.to_string()==*op && depth==0 && in_string.is_none() => {
                        let left = &s[..i]; let right = &s[i+op.len()..];
                        return Ok(Expr::Binary { op: op.to_string(), left: Box::new(parse_expr(left)?), right: Box::new(parse_expr(right)?) });
                    }
                    _=>{}
                }
            }
        }

        // Integer literal
        if let Ok(n) = s.parse::<i32>() { return Ok(Expr::Int(n)); }
        // Float
        if let Ok(f) = s.parse::<f64>() { return Ok(Expr::Float(f)); }
        // String
        if s.starts_with('"') && s.ends_with('"') && s.len()>=2 { return Ok(Expr::Str(s[1..s.len()-1].to_string())); }
        // Bool
        if s.eq_ignore_ascii_case("true") { return Ok(Expr::Bool(true)); }
        if s.eq_ignore_ascii_case("false") { return Ok(Expr::Bool(false)); }
        // None
        if s.eq_ignore_ascii_case("none") || s.eq_ignore_ascii_case("None") { return Ok(Expr::None); }

        // Identifier fallback
        Ok(Expr::Ident(s.to_string()))
    }

    let (stmts, _consumed) = parse_block(&lines, 0, 0)?;
    Ok(Program { stmts })
}
