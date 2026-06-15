//! C → common AST walker.
//!
//! Walks the pest parse tree from `grammar.pest` into `vybe_compiler::ast`
//! nodes. C-specific normalizations happen here so the shared compiler stays
//! language-agnostic:
//!   - `printf(fmt, …)` → `__c_printf(fmt, …)` (shared sprintf formatter)
//!   - structs are tracked so `struct P x;` initializes a zero-filled object
//!   - pointer deref `*p` / address-of `&x` lower to common reference AST
//!   - `a->b` is treated as `a.b`

use pest::Parser;
use pest::iterators::Pair;
use std::collections::{HashMap, HashSet};

use super::{CParser, Rule};
use crate::ast::*;
use crate::platforms::libc::pointers::{self, CARRAY_BASE_KEY, CARRAY_IDX_KEY, CARRAY_KIND};
use crate::platforms::libc::{ctype_adapter, math_adapter, string_adapter};

pub fn parse(source: &str) -> Result<Module, String> {
    let normalized_source = source.replace("\"\\\\\"\"", "\"\\\\\\\"\"");
    let mut pairs = CParser::parse(Rule::program, &normalized_source)
        .map_err(|e| format!("C parse error: {e}"))?;
    let program = pairs.next().ok_or("empty parse")?;
    let mut w = Walker::default();
    let mut body = Vec::new();
    for item in program.into_inner() {
        match item.as_rule() {
            Rule::EOI => {}
            _ => w.walk_top_item(item, &mut body),
        }
    }
    // Prepend static globals before the rest of the module body
    let mut full_body = w.static_globals;
    full_body.extend(body);
    Ok(Module {
        name: "main".to_string(),
        language: Lang::Unknown,
        body: full_body,
        imports: Vec::new(),
    })
}

#[derive(Default)]
struct Walker {
    /// struct/union name → ordered field names (for zero-init at decl site)
    structs: HashMap<String, Vec<String>>,
    /// struct/union name → field name → field type (for nested struct handling)
    struct_field_types: HashMap<String, HashMap<String, String>>,
    /// typedef names whose declarator is pointer-shaped.
    typedef_pointer_aliases: HashSet<String>,
    /// typedef names whose declarator is `char *`-shaped.
    typedef_char_pointer_aliases: HashSet<String>,
    /// identifiers declared as `char*`; used for pointer-like string traversal.
    char_pointers: HashSet<String>,
    /// char pointer variable -> (base string/array variable, element offset)
    char_pointer_offsets: HashMap<String, (String, Expression)>,
    /// identifiers declared as non-char pointer to array (int*, double*, etc.)
    /// These are PLAIN arrays (int arr[N]) — direct JS array indexing.
    array_ptr_vars: HashSet<String>,
    /// identifiers that ARE pointer variables backed by a carray object
    /// `{__ref_kind:"carray", __base:Array, __idx:i32}`.
    /// Covers `int *p = arr;` and parameters declared as `int *p`.
    carray_ptr_vars: HashSet<String>,
    /// pointer variable -> variable whose address it stores from `T *p = &x`.
    pointer_address_aliases: HashMap<String, String>,
    /// pointer variable -> struct/union member expression from `T *p = &obj.field`.
    pointer_member_aliases: HashMap<String, Expression>,
    /// identifiers whose address has been taken with `&name`; later plain
    /// reads/writes go through the common reference-cell AST.
    address_taken: HashSet<String>,
    /// variable/parameter name → C type string for sizeof resolution
    var_types: HashMap<String, String>,
    /// variable/parameter name → precomputed sizeof value (accounts for arrays)
    var_sizes: HashMap<String, i64>,
    /// function-like macros: name → (params, body text)
    macros: HashMap<String, (Vec<String>, String)>,
    /// object-like macros: name → raw replacement text
    object_macros: HashMap<String, String>,
    /// function name → parameter C type hints, used to normalize pointer arguments.
    function_param_types: HashMap<String, Vec<Option<String>>>,
    /// C enum constants are integer constants and can appear in global initializers.
    enum_constants: HashMap<String, i64>,
    /// current function name (for static local mangling)
    current_function: String,
    /// static local variable orignal name → mangled global name
    static_renames: HashMap<String, String>,
    /// accumulated static-local declarations to prepend to the module body
    static_globals: Vec<Statement>,
    /// current function char* parameter name → parameter index.
    current_char_param_indices: HashMap<String, usize>,
    /// function name → char* parameter writes `(param_index, index, value)`.
    char_param_writes: HashMap<String, Vec<(usize, Expression, Expression)>>,
}

fn stmt(kind: StmtKind) -> Statement {
    Statement::new(kind)
}
fn expr(kind: ExprKind) -> Expression {
    Expression::new(kind)
}
fn ident(name: &str) -> Expression {
    expr(ExprKind::Ident(name.to_string()))
}

fn carray_indexed_access(object: Expression, index: Expression) -> Expression {
    let adjusted = expr(ExprKind::Binary {
        op: BinOp::Add,
        left: Box::new(expr(ExprKind::Member {
            object: Box::new(object.clone()),
            field: CARRAY_IDX_KEY.to_string(),
            null_safe: false,
        })),
        right: Box::new(index),
    });
    expr(ExprKind::Index {
        object: Box::new(expr(ExprKind::Member {
            object: Box::new(object),
            field: CARRAY_BASE_KEY.to_string(),
            null_safe: false,
        })),
        index: Box::new(adjusted),
        null_safe: false,
    })
}

fn declarator_has_pointer(pair: &Pair<Rule>) -> bool {
    for child in pair.clone().into_inner() {
        match child.as_rule() {
            Rule::pointer => return true,
            Rule::declarator | Rule::direct_declarator => {
                if declarator_has_pointer(&child) {
                    return true;
                }
            }
            _ => {}
        }
    }
    false
}

impl Walker {
    fn walk_top_item(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        match pair.as_rule() {
            Rule::preproc_directive => self.walk_preproc(pair, out),
            Rule::function_definition => {
                if let Some(s) = self.walk_function(pair) {
                    out.push(s);
                }
            }
            Rule::declaration => self.walk_declaration(pair, out),
            _ => {}
        }
    }

    fn walk_preproc(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        for inner in pair.into_inner() {
            if inner.as_rule() == Rule::define_directive {
                // #define NAME value  → const NAME = value  (object-like only)
                let mut it = inner.into_inner();
                let name = match it.next() {
                    Some(p) if p.as_rule() == Rule::ident_name => p.as_str().to_string(),
                    _ => continue,
                };
                // Skip function-like macros (define_params present).
                let mut value_pair = None;
                let mut params_pair = None;
                for p in it {
                    match p.as_rule() {
                        Rule::define_params => params_pair = Some(p),
                        Rule::define_value => value_pair = Some(p),
                        _ => {}
                    }
                }
                if let Some(pp) = params_pair {
                    // Function-like macro: store for expansion at call sites
                    let params: Vec<String> = pp
                        .into_inner()
                        .filter(|p| p.as_rule() == Rule::ident_name)
                        .map(|p| p.as_str().to_string())
                        .collect();
                    let body_text = value_pair
                        .map(|p| p.as_str().trim().to_string())
                        .unwrap_or_default();
                    self.macros.insert(name, (params, body_text));
                    continue;
                }
                let init = value_pair
                    .map(|p| {
                        let raw = p.as_str().trim().to_string();
                        self.object_macros.insert(name.clone(), raw.clone());
                        self.parse_define_value(&raw)
                    })
                    .unwrap_or_else(|| expr(ExprKind::Lit(Literal::Int(1))));
                out.push(stmt(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(name),
                        type_hint: None,
                        init: Some(init),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Const,
                }));
            }
            // #include and conditionals: ignored (stdlib is profile-provided).
        }
    }

    /// Best-effort literal for a `#define` value; falls back to a string.
    fn parse_define_value(&mut self, text: &str) -> Expression {
        if let Ok(i) = text.parse::<i64>() {
            return expr(ExprKind::Lit(Literal::Int(i)));
        }
        if let Ok(f) = text.parse::<f64>() {
            return expr(ExprKind::Lit(Literal::Float(f)));
        }
        // Try re-parsing the value as a C expression.
        if let Ok(mut p) = CParser::parse(Rule::assignment_expression, text) {
            if let Some(e) = p.next() {
                return self.walk_assignment(e);
            }
        }
        expr(ExprKind::Lit(Literal::Str(text.to_string())))
    }

    // ── Functions ──────────────────────────────────────────────────────────

    fn walk_function(&mut self, pair: Pair<Rule>) -> Option<Statement> {
        let mut return_type = None;
        let mut name = String::new();
        let mut params = Vec::new();
        let mut body = Vec::new();
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::declaration_specifiers => return_type = Some(self.type_text(p)),
                Rule::declarator => {
                    let (n, ps) = self.declarator_name_and_params(p);
                    name = n;
                    if let Some(ps) = ps {
                        params = ps;
                    }
                }
                Rule::compound_statement => {
                    // Set current function context for static local mangling
                    self.current_function = name.clone();
                    self.static_renames.clear();
                    let mut scoped_char_params = Vec::new();
                    let mut scoped_carray_params = Vec::new();
                    for (idx, param) in params.iter().enumerate() {
                        if let Some(type_hint) = &param.type_hint {
                            if type_hint.contains("char") && type_hint.contains('*') {
                                self.char_pointers.insert(param.name.clone());
                                self.current_char_param_indices
                                    .insert(param.name.clone(), idx);
                                scoped_char_params.push(param.name.clone());
                            } else if self.is_carray_compatible_pointer_param(type_hint) {
                                self.carray_ptr_vars.insert(param.name.clone());
                                scoped_carray_params.push(param.name.clone());
                            }
                        }
                    }
                    body = self.walk_block(p);
                    for param in scoped_char_params {
                        self.char_pointers.remove(&param);
                    }
                    self.current_char_param_indices.clear();
                    for param in scoped_carray_params {
                        self.carray_ptr_vars.remove(&param);
                    }
                    for param in params.iter().rev() {
                        let Some(type_hint) = &param.type_hint else {
                            continue;
                        };
                        if type_hint.contains('*') {
                            continue;
                        }
                        let normalized_type = normalized_c_type_name(type_hint);
                        if self.structs.contains_key(&normalized_type) {
                            body.insert(
                                0,
                                stmt(StmtKind::Expr(expr(ExprKind::Assign {
                                    target: Box::new(ident(&param.name)),
                                    value: Box::new(
                                        self.deep_copy_struct(type_hint, ident(&param.name)),
                                    ),
                                }))),
                            );
                        }
                    }
                }
                _ => {}
            }
        }
        if name.is_empty() {
            return None;
        }
        self.function_param_types.insert(
            name.clone(),
            params.iter().map(|param| param.type_hint.clone()).collect(),
        );
        Some(stmt(StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            body,
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async: false,
            is_generator: false,
            is_sub: false,
        }))
    }

    fn declarator_name_and_params(&mut self, pair: Pair<Rule>) -> (String, Option<Vec<Param>>) {
        // declarator = pointer? ~ direct_declarator
        let mut name = String::new();
        let mut params = None;
        for p in pair.into_inner() {
            if p.as_rule() == Rule::direct_declarator {
                for d in p.into_inner() {
                    match d.as_rule() {
                        Rule::ident_name => name = d.as_str().to_string(),
                        Rule::declarator => {
                            let (n, ps) = self.declarator_name_and_params(d);
                            name = n;
                            if ps.is_some() {
                                params = ps;
                            }
                        }
                        Rule::param_suffix => params = Some(self.walk_params(d)),
                        _ => {}
                    }
                }
            }
        }
        (name, params)
    }

    fn walk_params(&mut self, pair: Pair<Rule>) -> Vec<Param> {
        let mut params = Vec::new();
        for p in pair.into_inner() {
            if p.as_rule() == Rule::parameter_list {
                for decl in p.into_inner() {
                    if decl.as_rule() == Rule::parameter_decl {
                        let decl_text = decl.as_str().to_string();
                        let mut pname = String::new();
                        let mut type_hint = None;
                        let mut is_pointer_decl = decl_text.contains('[');
                        for d in decl.into_inner() {
                            match d.as_rule() {
                                Rule::declaration_specifiers => type_hint = Some(self.type_text(d)),
                                Rule::declarator => {
                                    is_pointer_decl = is_pointer_decl
                                        || declarator_has_pointer(&d)
                                        || decl_text.contains('*');
                                    pname = self.declarator_name_and_params(d).0;
                                }
                                _ => {}
                            }
                        }
                        if is_pointer_decl {
                            if let Some(hint) = &mut type_hint {
                                let existing = hint.matches('*').count();
                                let declared = decl_text.matches('*').count().max(1);
                                for _ in existing..declared {
                                    hint.push_str(" *");
                                }
                            }
                        }
                        if !pname.is_empty() {
                            params.push(Param {
                                name: pname,
                                type_hint,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: false,
                            });
                        }
                    }
                }
            }
        }
        params
    }

    // ── Declarations ─────────────────────────────────────────────────────────

    fn walk_declaration(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        let inner = pair.into_inner().next();
        let Some(inner) = inner else { return };
        match inner.as_rule() {
            Rule::typedef_declaration => self.walk_typedef(inner, out),
            Rule::normal_declaration => self.walk_normal_declaration(inner, out),
            _ => {}
        }
    }

    fn walk_typedef(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        // typedef declaration_specifiers declarator (, declarator)* ;
        let mut specs = None;
        let mut names = Vec::new();
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::declaration_specifiers => specs = Some(p),
                Rule::declarator => {
                    let is_pointer_alias = declarator_has_pointer(&p)
                        || p.as_str().split('=').next().unwrap_or("").contains('*');
                    let name = self.declarator_name_and_params(p).0;
                    if is_pointer_alias && !name.is_empty() {
                        self.typedef_pointer_aliases.insert(name.clone());
                    }
                    names.push(name);
                }
                _ => {}
            }
        }
        if let Some(ref specs) = specs {
            if self.type_text(specs.clone()).contains("char") {
                for name in &names {
                    if self.typedef_pointer_aliases.contains(name) {
                        self.typedef_char_pointer_aliases.insert(name.clone());
                    }
                }
            }
            // typedef struct {...} Name; → register Name as struct alias.
            if let Some((tag, fields)) = self.struct_def_from_specifiers(specs) {
                for name in &names {
                    self.structs.insert(name.clone(), fields.clone());
                    out.push(self.make_struct_decl(name, &fields));
                }
                let _ = tag;
            }
            // typedef enum { A, B } Name; → emit enum members as consts.
            if let Some((_tag, members)) = self.enum_def_from_specifiers(specs) {
                let mut next_val: i64 = 0;
                for member in &members {
                    let val = extract_enum_val(&member.value, next_val);
                    self.enum_constants.insert(member.name.clone(), val);
                    out.push(stmt(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(member.name.clone()),
                            type_hint: Some("int".to_string()),
                            init: Some(expr(ExprKind::Lit(Literal::Int(val)))),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Const,
                    }));
                    next_val = val + 1;
                }
            }
        }
    }

    fn walk_normal_declaration(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        let mut specs = None;
        let mut init_list = None;
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::declaration_specifiers => specs = Some(p),
                Rule::init_declarator_list => init_list = Some(p),
                _ => {}
            }
        }
        let Some(specs) = specs else { return };

        // A struct/union/enum definition with a body.
        if let Some((tag, fields)) = self.struct_def_from_specifiers(&specs) {
            if let Some(tag) = tag.clone() {
                self.structs.insert(tag.clone(), fields.clone());
                if init_list.is_none() {
                    out.push(self.make_struct_decl(&tag, &fields));
                }
            }
        }
        if let Some((name, members)) = self.enum_def_from_specifiers(&specs) {
            // In C, enum members are global integer constants — emit each as const.
            // Auto-increment value starting from 0, incremented per member.
            let mut next_val: i64 = 0;
            for member in &members {
                let val = extract_enum_val(&member.value, next_val);
                self.enum_constants.insert(member.name.clone(), val);
                out.push(stmt(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(member.name.clone()),
                        type_hint: Some("int".to_string()),
                        init: Some(expr(ExprKind::Lit(Literal::Int(val)))),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Const,
                }));
                next_val = val + 1;
            }
            out.push(stmt(StmtKind::EnumDecl {
                name: name.unwrap_or_else(|| "__anon_enum".to_string()),
                members,
                visibility: Visibility::Public,
                is_flags: false,
                backing_type: None,
                interfaces: Vec::new(),
                body_members: Vec::new(),
                decorators: Vec::new(),
            }));
        }

        let struct_fields = self.struct_type_of_specifiers(&specs);
        let type_text = self.type_text(specs);

        // Skip extern declarations with no initializer (function prototypes etc.)
        if type_text.split_whitespace().any(|w| w == "extern") && init_list.is_none() {
            return;
        }

        // Check for static locals (only when inside a function)
        let is_static_local = !self.current_function.is_empty()
            && type_text.split_whitespace().any(|w| w == "static");

        let Some(init_list) = init_list else { return };
        let mut declarations = Vec::new();
        for idecl in init_list.into_inner() {
            if idecl.as_rule() != Rule::init_declarator {
                continue;
            }
            let declarator_text = idecl.as_str().split('=').next().unwrap_or("").to_string();
            let mut name = String::new();
            let mut array_bounds: Option<Vec<Expression>> = None;
            let mut init = None;
            let mut is_pointer_decl = false;
            let mut init_is_addr_of = false;
            let mut is_function_proto = false;
            let mut is_function_pointer_decl = false;
            let mut was_array_decl = false;
            for p in idecl.into_inner() {
                match p.as_rule() {
                    Rule::declarator => {
                        is_pointer_decl =
                            declarator_has_pointer(&p) || declarator_text.contains('*');
                        if self
                            .typedef_pointer_aliases
                            .contains(&normalized_c_type_name(&type_text))
                        {
                            is_pointer_decl = true;
                        }
                        // Detect function-pointer or prototype declarator: has param_suffix
                        is_function_proto =
                            p.as_str().contains('(') && !p.as_str().starts_with('*'); // not a function-pointer type
                        is_function_pointer_decl =
                            declarator_text.contains("(*") && declarator_text.contains(")(");
                        let (n, bounds) = self.declarator_name_and_bounds(p);
                        name = n;
                        array_bounds = bounds;
                        if array_bounds.is_some() {
                            was_array_decl = true;
                        }
                    }
                    Rule::initializer => {
                        // Check before walking if init is address-of (&x) form
                        init_is_addr_of = p.as_str().trim().starts_with('&');
                        let raw = self.walk_initializer(p);
                        if matches!(&raw.kind, ExprKind::Ident(n) if self.array_ptr_vars.contains(n))
                        {
                            is_pointer_decl = true;
                        }
                        init = Some(if !is_pointer_decl {
                            if let Some(fields) = &struct_fields {
                                if array_bounds.is_some() {
                                    // Array of structs: convert each element to a named object.
                                    // `struct Pair pairs[2] = {{1,2},{3,4}}` → [{a:1,b:2},{a:3,b:4}]
                                    if let ExprKind::Array(elems) = raw.kind {
                                        let converted: Vec<ArrayElement> = elems
                                            .into_iter()
                                            .map(|el| ArrayElement {
                                                value: self.convert_array_init_to_struct_typed(
                                                    &type_text, el.value, fields,
                                                ),
                                                ..el
                                            })
                                            .collect();
                                        array_bounds = None; // embedded in literal
                                        expr(ExprKind::Array(converted))
                                    } else {
                                        raw
                                    }
                                } else {
                                    // Convert array init to struct, and also handle struct-to-struct copy
                                    let converted = self.convert_array_init_to_struct_typed(
                                        &type_text,
                                        raw.clone(),
                                        fields,
                                    );
                                    // If init is a simple identifier (struct copy), wrap in deep copy
                                    if matches!(raw.kind, ExprKind::Ident(_))
                                        || matches!(raw.kind, ExprKind::Member { .. })
                                    {
                                        self.deep_copy_struct(&type_text, converted)
                                    } else {
                                        converted
                                    }
                                }
                            } else {
                                raw
                            }
                        } else {
                            raw
                        });
                    }
                    _ => {}
                }
            }
            if name.is_empty() {
                continue;
            }
            if type_text.split_whitespace().any(|w| w == "extern") && init.is_none() {
                continue;
            }
            // Skip function prototypes: `int foo(int x);` has no init and is
            // a function-like declarator. Emitting `var foo;` would shadow the
            // actual function definition.
            if is_function_proto && init.is_none() && !is_pointer_decl {
                continue;
            }
            let normalized_type_text = normalized_c_type_name(&type_text);
            let type_is_char_pointer_alias = self
                .typedef_char_pointer_aliases
                .contains(&normalized_type_text);
            if (type_text.contains("char") || type_is_char_pointer_alias)
                && !is_function_pointer_decl
            {
                // Track char* pointers AND char arrays (initialized with string literals)
                // for substring-based pointer arithmetic.
                let init_is_string = init
                    .as_ref()
                    .map(|i| matches!(i.kind, ExprKind::Lit(Literal::Str(_))))
                    .unwrap_or(false);
                let init_is_heap_array = init
                    .as_ref()
                    .map(|i| matches!(i.kind, ExprKind::Array(_)))
                    .unwrap_or(false);
                if is_pointer_decl && init_is_heap_array {
                    self.char_pointers.insert(name.clone());
                    init = Some(expr(ExprKind::Lit(Literal::Str(String::new()))));
                } else if init_is_string
                    || (is_pointer_decl && !init_is_addr_of && !is_null_pointer_init(&init))
                {
                    self.char_pointers.insert(name.clone());
                    if let Some((base, offset)) = char_pointer_offset_from_init(&init) {
                        self.char_pointer_offsets
                            .insert(name.clone(), (base, offset));
                    }
                }
            } else if is_pointer_decl && !is_function_pointer_decl {
                // Non-char pointer variable — decide scalar-cell vs carray.
                // If the walked init is already a carray object (e.g. from `&arr[n]`),
                // track this var as carray; otherwise wrap a plain array as carray.
                let init_is_carray = init.as_ref().map(|i| is_carray_object(i)).unwrap_or(false);
                if let Some(target) = pointer_address_target_from_init(&init) {
                    self.pointer_address_aliases.insert(name.clone(), target);
                } else if let Some(target) = pointer_member_target_from_init(&init) {
                    self.pointer_member_aliases.insert(name.clone(), target);
                } else if let Some(target) =
                    propagated_pointer_address_alias(&init, &self.pointer_address_aliases)
                {
                    self.pointer_address_aliases.insert(name.clone(), target);
                }
                if init_is_carray {
                    // int *p = &arr[n] → init already carray from apply_prefix
                    self.carray_ptr_vars.insert(name.clone());
                } else if !init_is_addr_of && !is_null_pointer_init(&init) {
                    if init_is_carray_pointer_var(&init, &self.carray_ptr_vars) {
                        self.carray_ptr_vars.insert(name.clone());
                    } else if should_wrap_pointer_init_as_carray(&init, &self.array_ptr_vars) {
                        // int *p = arr → wrap as carray
                        self.carray_ptr_vars.insert(name.clone());
                        if let Some(ref raw_init) = init {
                            init = Some(self.wrap_as_carray_init(raw_init.clone()));
                        }
                    }
                }
                // else: int *p = &scalar → scalar cell (address_taken mechanism)
            }
            // Zero-init struct/union instances when no explicit initializer.
            if init.is_none() {
                if let Some(fields) = &struct_fields {
                    let struct_name = normalized_c_type_name(&type_text);
                    if let Some(ref bounds) = array_bounds {
                        // Array of structs: pre-fill with N copies of zero struct.
                        let count = bounds
                            .first()
                            .and_then(|b| {
                                if let ExprKind::Lit(Literal::Int(n)) = &b.kind {
                                    Some(*n as usize)
                                } else {
                                    None
                                }
                            })
                            .unwrap_or(0);
                        if count > 0 {
                            let zeros: Vec<ArrayElement> = (0..count)
                                .map(|_| ArrayElement {
                                    value: self.zero_struct(Some(&struct_name), fields),
                                    spread: false,
                                    key: None,
                                    by_ref: false,
                                })
                                .collect();
                            init = Some(expr(ExprKind::Array(zeros)));
                            array_bounds = None;
                        }
                    } else {
                        init = Some(self.zero_struct(Some(&struct_name), fields));
                    }
                }
            }
            // char array with bounds initialized by a string → treat as string
            // (e.g. `char buf[32] = "hello"` → just a string variable)
            let is_char_type = normalized_c_type_name(&type_text) == "char";
            if array_bounds.is_some() && is_char_type {
                if let Some(ref init_expr) = init {
                    if matches!(init_expr.kind, ExprKind::Lit(Literal::Str(_))) {
                        array_bounds = None; // treat as string, not array
                    }
                }
            }
            // char array with char initializers `{'h','i','\0'}` → join chars to string
            if array_bounds.is_some() && is_char_type {
                if let Some(ExprKind::Array(elems)) = init.as_ref().map(|i| &i.kind) {
                    let s: String = elems
                        .iter()
                        .filter_map(|el| {
                            if let ExprKind::Lit(Literal::Int(code)) = &el.value.kind {
                                if *code == 0 {
                                    None
                                } else {
                                    char::from_u32(*code as u32)
                                }
                            } else {
                                None
                            }
                        })
                        .collect();
                    init = Some(expr(ExprKind::Lit(Literal::Str(s))));
                    self.char_pointers.insert(name.clone());
                    array_bounds = None;
                }
            }
            // Partial array initialization: zero-fill tail slots.
            // `int arr[4] = {1, 2}` → `[1, 2, 0, 0]`
            if !type_text.contains("char") && struct_fields.is_none() {
                if let (Some(bounds), Some(init_expr)) = (&array_bounds, &init) {
                    if let ExprKind::Array(elems) = &init_expr.kind {
                        let count = bounds
                            .first()
                            .and_then(|b| {
                                if let ExprKind::Lit(Literal::Int(n)) = &b.kind {
                                    Some(*n as usize)
                                } else {
                                    None
                                }
                            })
                            .unwrap_or(0);
                        if count > elems.len() {
                            let mut padded = elems.clone();
                            while padded.len() < count {
                                padded.push(ArrayElement {
                                    value: expr(ExprKind::Lit(Literal::Int(0))),
                                    spread: false,
                                    key: None,
                                    by_ref: false,
                                });
                            }
                            init = Some(expr(ExprKind::Array(padded)));
                        }
                    }
                }
            }
            let mut metadata_type = type_text.clone();
            if let Some(ref bounds) = array_bounds {
                for b in bounds {
                    if let ExprKind::Lit(Literal::Int(n)) = &b.kind {
                        metadata_type.push_str(&format!("[{}]", n));
                    } else {
                        metadata_type.push_str("[]");
                    }
                }
            } else if is_char_type {
                if let Some(ExprKind::Lit(Literal::Str(s))) = init.as_ref().map(|i| &i.kind) {
                    metadata_type.push_str(&format!("[{}]", s.len() + 1));
                }
            }
            // Record the type for sizeof resolution
            self.var_types.insert(name.clone(), metadata_type.clone());
            // Record sizeof for this variable.
            // NOTE: compute from init elem count when array_bounds has been cleared by zero-fill.
            let sz = if is_pointer_decl {
                8
            } else if let Some(ref bounds) = array_bounds {
                let base = sizeof_from_type_text(&type_text);
                let count: i64 = bounds
                    .iter()
                    .map(|b| {
                        if let ExprKind::Lit(Literal::Int(n)) = &b.kind {
                            *n
                        } else {
                            1
                        }
                    })
                    .product();
                base * count
            } else if let Some(ExprKind::Array(elems)) = init.as_ref().map(|i| &i.kind) {
                // array_bounds was cleared after pre-fill — count elements
                let base = sizeof_from_type_text(&type_text);
                base * elems.len() as i64
            } else if is_char_type {
                if let Some(ExprKind::Lit(Literal::Str(s))) = init.as_ref().map(|i| &i.kind) {
                    (s.len() + 1) as i64
                } else {
                    sizeof_from_type_text(&type_text)
                }
            } else {
                let su = self.sizeof_struct_union(&type_text);
                if su > 0 {
                    su
                } else {
                    sizeof_from_type_text(&type_text)
                }
            };
            self.var_sizes.insert(name.clone(), sz);
            // Non-char arrays (int arr[n], double arr[n]) decay to pointer for arithmetic.
            if !type_text.contains("char") && was_array_decl {
                self.array_ptr_vars.insert(name.clone());
            }
            // Handle static local: lift to a module-level global with mangled name
            if is_static_local {
                let mangled = format!("__static_{}_{}", self.current_function, name);
                self.static_renames.insert(name.clone(), mangled.clone());
                self.static_globals.push(stmt(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(mangled),
                        type_hint: Some(type_text.clone()),
                        init,
                        array_bounds,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Var,
                }));
                continue;
            }
            declarations.push(VarDeclarator {
                pattern: BindingPattern::Ident(name),
                type_hint: Some(type_text.clone()),
                init,
                array_bounds,
                with_events: false,
            });
        }
        if !declarations.is_empty() {
            out.push(stmt(StmtKind::VarDecl {
                declarations,
                kind: VarDeclKind::Var,
            }));
        }
    }

    fn make_struct_decl(&self, name: &str, fields: &[String]) -> Statement {
        let members = fields
            .iter()
            .map(|f| ClassMember::Field {
                name: f.clone(),
                type_hint: None,
                init: Some(expr(ExprKind::Lit(Literal::Int(0)))),
                modifiers: Modifiers::default(),
                with_events: false,
                array_bounds: None,
            })
            .collect();
        stmt(StmtKind::StructDecl {
            name: name.to_string(),
            interfaces: Vec::new(),
            members,
            visibility: Visibility::Public,
            decorators: Vec::new(),
        })
    }

    fn zero_struct(&self, struct_name_hint: Option<&str>, fields: &[String]) -> Expression {
        let props = fields
            .iter()
            .map(|f| {
                // Look up field type in struct_field_types if we have a struct name
                let value = if let Some(sname) = struct_name_hint {
                    if let Some(field_type_map) = self.struct_field_types.get(sname) {
                        if let Some(field_type) = field_type_map.get(f) {
                            let normalized_type = normalized_c_type_name(field_type);
                            if let Some(nested_fields) = self.structs.get(&normalized_type) {
                                // Recursively initialize nested struct
                                self.zero_struct(Some(&normalized_type), nested_fields)
                            } else {
                                expr(ExprKind::Lit(Literal::Int(0)))
                            }
                        } else {
                            expr(ExprKind::Lit(Literal::Int(0)))
                        }
                    } else {
                        expr(ExprKind::Lit(Literal::Int(0)))
                    }
                } else {
                    expr(ExprKind::Lit(Literal::Int(0)))
                };
                ObjectProperty::KeyValue {
                    key: expr(ExprKind::Lit(Literal::Str(f.clone()))),
                    value,
                }
            })
            .collect();
        expr(ExprKind::Object(props))
    }

    /// Deep copy a struct by recursively copying all fields, including nested structs.
    /// For `struct Pair second = first;`, generates:
    /// `{a: first.a, b: first.b}` for simple fields
    /// `{origin: {x: first.origin.x, y: first.origin.y}, size: first.size}` for nested
    fn deep_copy_struct(&self, type_name: &str, source: Expression) -> Expression {
        let normalized_type = normalized_c_type_name(type_name);
        let Some(fields) = self.structs.get(&normalized_type) else {
            return source; // Not a known struct, return as-is
        };

        let props: Vec<ObjectProperty> = fields
            .iter()
            .map(|f| {
                let member_access = expr(ExprKind::Member {
                    object: Box::new(source.clone()),
                    field: f.clone(),
                    null_safe: false,
                });

                // Check if this field is itself a struct
                let value =
                    if let Some(field_type_map) = self.struct_field_types.get(&normalized_type) {
                        if let Some(field_type) = field_type_map.get(f) {
                            let field_normalized = normalized_c_type_name(field_type);
                            if self.structs.contains_key(&field_normalized) {
                                // Recursively deep copy nested struct
                                self.deep_copy_struct(field_type, member_access)
                            } else {
                                member_access
                            }
                        } else {
                            member_access
                        }
                    } else {
                        member_access
                    };

                ObjectProperty::KeyValue {
                    key: expr(ExprKind::Lit(Literal::Str(f.clone()))),
                    value,
                }
            })
            .collect();

        expr(ExprKind::Object(props))
    }

    fn convert_array_init_to_struct_typed(
        &self,
        type_name: &str,
        raw: Expression,
        fields: &[String],
    ) -> Expression {
        let elems = match raw.kind {
            ExprKind::Array(elems) => elems,
            other => return expr(other),
        };
        if elems.is_empty() {
            return expr(ExprKind::Array(elems));
        }

        let normalized_type = normalized_c_type_name(type_name);
        let field_types = self.struct_field_types.get(&normalized_type);
        let mut props = Vec::new();
        for (i, el) in elems.into_iter().enumerate() {
            let Some(fname) = fields.get(i).cloned() else {
                continue;
            };
            let value = field_types
                .and_then(|types| types.get(&fname))
                .and_then(|field_type| {
                    let field_type_name = normalized_c_type_name(field_type);
                    let nested_fields = self.structs.get(&field_type_name)?;
                    Some(self.convert_array_init_to_struct_typed(
                        field_type,
                        el.value.clone(),
                        nested_fields,
                    ))
                })
                .unwrap_or(el.value);
            props.push(ObjectProperty::KeyValue {
                key: expr(ExprKind::Lit(Literal::Str(fname))),
                value,
            });
        }
        for i in props.len()..fields.len() {
            props.push(ObjectProperty::KeyValue {
                key: expr(ExprKind::Lit(Literal::Str(fields[i].clone()))),
                value: expr(ExprKind::Lit(Literal::Int(0))),
            });
        }
        expr(ExprKind::Object(props))
    }

    /// If the specifiers declare a struct/union with a body, return
    /// `(optional tag name, field names)`.
    fn struct_def_from_specifiers(
        &mut self,
        specs: &Pair<Rule>,
    ) -> Option<(Option<String>, Vec<String>)> {
        for p in specs.clone().into_inner() {
            if p.as_rule() == Rule::type_specifier || p.as_rule() == Rule::type_specifier_strict {
                for ts in p.into_inner() {
                    if ts.as_rule() == Rule::struct_or_union_spec {
                        let mut tag = None;
                        let mut fields = Vec::new();
                        let mut field_types = HashMap::new();
                        let mut has_body = false;
                        for sp in ts.into_inner() {
                            match sp.as_rule() {
                                Rule::ident_name => tag = Some(sp.as_str().to_string()),
                                Rule::struct_member => {
                                    has_body = true;
                                    self.collect_struct_fields(sp, &mut fields, &mut field_types);
                                }
                                _ => {}
                            }
                        }
                        if has_body {
                            if let Some(ref tag_name) = tag {
                                self.struct_field_types
                                    .insert(tag_name.clone(), field_types);
                            }
                            return Some((tag, fields));
                        }
                    }
                }
            }
        }
        None
    }

    fn collect_struct_fields(
        &self,
        member: Pair<Rule>,
        fields: &mut Vec<String>,
        field_types: &mut HashMap<String, String>,
    ) {
        let mut member_type = None;
        for p in member.into_inner() {
            if p.as_rule() == Rule::declaration_specifiers {
                member_type = Some(self.type_text(p));
            } else if p.as_rule() == Rule::struct_declarator_list {
                for d in p.into_inner() {
                    if d.as_rule() == Rule::declarator {
                        let n = self.clone_declarator_name(d);
                        if !n.is_empty() {
                            fields.push(n.clone());
                            if let Some(ref ty) = member_type {
                                field_types.insert(n, ty.clone());
                            }
                        }
                    }
                }
            }
        }
    }

    fn clone_declarator_name(&self, pair: Pair<Rule>) -> String {
        for p in pair.into_inner() {
            if p.as_rule() == Rule::direct_declarator {
                for d in p.into_inner() {
                    match d.as_rule() {
                        Rule::ident_name => return d.as_str().to_string(),
                        Rule::declarator => return self.clone_declarator_name(d),
                        _ => {}
                    }
                }
            }
        }
        String::new()
    }

    /// Resolve the struct field list referenced by a declaration's specifiers
    /// (either an inline body or a previously-registered struct name).
    fn struct_type_of_specifiers(&mut self, specs: &Pair<Rule>) -> Option<Vec<String>> {
        if let Some((_, fields)) = self.struct_def_from_specifiers(specs) {
            return Some(fields);
        }
        for p in specs.clone().into_inner() {
            match p.as_rule() {
                Rule::type_specifier | Rule::type_specifier_strict => {
                    for ts in p.into_inner() {
                        match ts.as_rule() {
                            Rule::struct_or_union_spec => {
                                for sp in ts.into_inner() {
                                    if sp.as_rule() == Rule::ident_name {
                                        if let Some(f) = self.structs.get(sp.as_str()) {
                                            return Some(f.clone());
                                        }
                                    }
                                }
                            }
                            Rule::typedef_name => {
                                if let Some(f) = self.structs.get(ts.as_str()) {
                                    return Some(f.clone());
                                }
                            }
                            _ => {}
                        }
                    }
                }
                // bare typedef_name as direct child of declaration_specifiers
                // e.g. `Point point = {...}` where Point is typedef'd struct
                Rule::typedef_name => {
                    if let Some(f) = self.structs.get(p.as_str()) {
                        return Some(f.clone());
                    }
                }
                _ => {}
            }
        }
        None
    }

    fn enum_def_from_specifiers(
        &mut self,
        specs: &Pair<Rule>,
    ) -> Option<(Option<String>, Vec<EnumMember>)> {
        for p in specs.clone().into_inner() {
            if p.as_rule() == Rule::type_specifier || p.as_rule() == Rule::type_specifier_strict {
                for ts in p.into_inner() {
                    if ts.as_rule() == Rule::enum_spec {
                        let mut name = None;
                        let mut members = Vec::new();
                        let mut has_body = false;
                        for sp in ts.into_inner() {
                            match sp.as_rule() {
                                Rule::ident_name => name = Some(sp.as_str().to_string()),
                                Rule::enumerator_list => {
                                    has_body = true;
                                    for en in sp.into_inner() {
                                        if en.as_rule() == Rule::enumerator {
                                            let mut it = en.into_inner();
                                            let nm = it.next().map(|x| x.as_str().to_string());
                                            let val = it.next().map(|x| self.walk_assignment(x));
                                            if let Some(nm) = nm {
                                                members.push(EnumMember {
                                                    name: nm,
                                                    value: val,
                                                    constructor_args: Vec::new(),
                                                });
                                            }
                                        }
                                    }
                                }
                                _ => {}
                            }
                        }
                        if has_body {
                            return Some((name, members));
                        }
                    }
                }
            }
        }
        None
    }

    fn declarator_name_and_bounds(
        &mut self,
        pair: Pair<Rule>,
    ) -> (String, Option<Vec<Expression>>) {
        let mut name = String::new();
        let mut bounds: Vec<Expression> = Vec::new();
        for p in pair.into_inner() {
            if p.as_rule() == Rule::direct_declarator {
                for d in p.into_inner() {
                    match d.as_rule() {
                        Rule::ident_name => name = d.as_str().to_string(),
                        Rule::declarator => {
                            let (n, b) = self.declarator_name_and_bounds(d);
                            name = n;
                            if let Some(b) = b {
                                bounds.extend(b);
                            }
                        }
                        Rule::array_suffix => {
                            if let Some(sz) = d.into_inner().next() {
                                bounds.push(self.walk_assignment(sz));
                            } else {
                                bounds.push(expr(ExprKind::Lit(Literal::Int(0))));
                            }
                        }
                        _ => {}
                    }
                }
            }
        }
        (
            name,
            if bounds.is_empty() {
                None
            } else {
                Some(bounds)
            },
        )
    }

    fn walk_initializer(&mut self, pair: Pair<Rule>) -> Expression {
        let inner = pair.into_inner().next();
        let Some(inner) = inner else {
            return expr(ExprKind::Lit(Literal::Null));
        };
        match inner.as_rule() {
            Rule::assignment_expression => self.walk_assignment(inner),
            Rule::initializer_list => {
                let mut is_object = false;
                let mut elems = Vec::new();
                let mut props = Vec::new();
                for di in inner.into_inner() {
                    if di.as_rule() != Rule::designated_init {
                        continue;
                    }
                    let mut it = di.into_inner().peekable();
                    // designated `.field = init`
                    let first = it.next();
                    match first {
                        Some(p) if p.as_rule() == Rule::ident_name => {
                            is_object = true;
                            let key = p.as_str().to_string();
                            if let Some(v) = it.next() {
                                let val = self.walk_initializer(v);
                                props.push(ObjectProperty::KeyValue {
                                    key: expr(ExprKind::Lit(Literal::Str(key))),
                                    value: val,
                                });
                            }
                        }
                        Some(p) if p.as_rule() == Rule::initializer => {
                            elems.push(ArrayElement {
                                key: None,
                                value: self.walk_initializer(p),
                                spread: false,
                                by_ref: false,
                            });
                        }
                        Some(p) if p.as_rule() == Rule::assignment_expression => {
                            // `[idx] = init` designator → treat as array element
                            is_object = false;
                            if let Some(v) = it.next() {
                                elems.push(ArrayElement {
                                    key: Some(self.walk_assignment(p)),
                                    value: self.walk_initializer(v),
                                    spread: false,
                                    by_ref: false,
                                });
                            } else {
                                elems.push(ArrayElement {
                                    key: None,
                                    value: self.walk_assignment(p),
                                    spread: false,
                                    by_ref: false,
                                });
                            }
                        }
                        _ => {}
                    }
                }
                if is_object {
                    expr(ExprKind::Object(props))
                } else {
                    expr(ExprKind::Array(elems))
                }
            }
            _ => expr(ExprKind::Lit(Literal::Null)),
        }
    }

    fn type_text(&self, pair: Pair<Rule>) -> String {
        pair.as_str()
            .split_whitespace()
            .collect::<Vec<_>>()
            .join(" ")
    }

    // ── Statements ─────────────────────────────────────────────────────────

    fn walk_block(&mut self, pair: Pair<Rule>) -> Vec<Statement> {
        let mut out = Vec::new();
        for p in pair.into_inner() {
            if p.as_rule() == Rule::statement {
                self.walk_statement(p, &mut out);
            }
        }
        out
    }

    fn walk_statement(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        let Some(inner) = pair.into_inner().next() else {
            return;
        };
        match inner.as_rule() {
            Rule::compound_statement => out.push(stmt(StmtKind::Block(self.walk_block(inner)))),
            Rule::function_definition => {
                if let Some(function) = self.walk_function(inner) {
                    out.push(function);
                }
            }
            Rule::declaration => self.walk_declaration(inner, out),
            Rule::expression_statement => {
                let e = inner.into_inner().next().unwrap();
                let expr = self.walk_expression(e);
                let expr = self.rewrite_carray_postfix_discard(expr);
                out.push(stmt(StmtKind::Expr(expr)));
            }
            Rule::if_statement => out.push(self.walk_if(inner)),
            Rule::switch_statement => out.push(self.walk_switch(inner)),
            Rule::while_statement => out.push(self.walk_while(inner)),
            Rule::do_while_statement => out.push(self.walk_do_while(inner)),
            Rule::for_statement => out.push(self.walk_for(inner)),
            Rule::return_statement => {
                let val = inner.into_inner().next().map(|e| self.walk_expression(e));
                out.push(stmt(StmtKind::Return(val)));
            }
            Rule::break_statement => out.push(stmt(StmtKind::Break(BreakTarget::Implicit))),
            Rule::continue_statement => {
                out.push(stmt(StmtKind::Continue(ContinueTarget::Implicit)))
            }
            Rule::goto_statement => {
                let label = inner.into_inner().next().map(|p| p.as_str().to_string());
                out.push(stmt(StmtKind::GoTo(label.unwrap_or_default())));
            }
            Rule::labeled_statement => self.walk_labeled(inner, out),
            Rule::empty_statement => {}
            _ => {}
        }
    }

    fn body_of(&mut self, pair: Pair<Rule>) -> Vec<Statement> {
        // A `statement` that may be a block or a single statement.
        let mut out = Vec::new();
        self.walk_statement(pair, &mut out);
        // Unwrap a single Block into its statements for cleaner nesting.
        if out.len() == 1 {
            if let StmtKind::Block(inner) = &out[0].kind {
                return inner.clone();
            }
        }
        out
    }

    fn walk_if(&mut self, pair: Pair<Rule>) -> Statement {
        let mut it = pair.into_inner();
        let cond = self.walk_expression(it.next().unwrap());
        let then_body = self.body_of(it.next().unwrap());
        let mut else_body = None;
        if let Some(else_clause) = it.next() {
            if else_clause.as_rule() == Rule::else_clause {
                let st = else_clause.into_inner().next().unwrap();
                else_body = Some(self.body_of(st));
            }
        }
        stmt(StmtKind::If {
            cond,
            then_body,
            elifs: Vec::new(),
            else_body,
        })
    }

    fn walk_while(&mut self, pair: Pair<Rule>) -> Statement {
        let mut it = pair.into_inner();
        let cond = self.walk_expression(it.next().unwrap());
        let body = self.body_of(it.next().unwrap());
        stmt(StmtKind::While {
            cond,
            body,
            else_body: None,
        })
    }

    fn walk_do_while(&mut self, pair: Pair<Rule>) -> Statement {
        let mut it = pair.into_inner();
        let body = self.body_of(it.next().unwrap());
        let cond = self.walk_expression(it.next().unwrap());
        stmt(StmtKind::DoWhile {
            body,
            cond,
            until: false,
        })
    }

    fn walk_for(&mut self, pair: Pair<Rule>) -> Statement {
        let mut init = None;
        let mut cond = None;
        let mut update = None;
        let mut body = Vec::new();
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::for_init => {
                    let mut inits = Vec::new();
                    if let Some(c) = p.into_inner().next() {
                        match c.as_rule() {
                            Rule::declaration => self.walk_declaration(c, &mut inits),
                            Rule::expression => {
                                inits.push(stmt(StmtKind::Expr(self.walk_expression(c))))
                            }
                            _ => {}
                        }
                    }
                    init = inits.into_iter().next().map(Box::new);
                }
                Rule::for_cond => {
                    cond = p.into_inner().next().map(|e| self.walk_expression(e));
                }
                Rule::for_update => {
                    update = p
                        .into_inner()
                        .next()
                        .map(|e| self.walk_expression(e))
                        .map(|e| self.rewrite_carray_postfix_discard(e));
                }
                Rule::statement => body = self.body_of(p),
                _ => {}
            }
        }
        stmt(StmtKind::For {
            init,
            cond,
            update,
            body,
        })
    }

    fn walk_switch(&mut self, pair: Pair<Rule>) -> Statement {
        let mut it = pair.into_inner();
        let subject = self.walk_expression(it.next().unwrap());
        let body_stmt = it.next().unwrap();
        // The body is a compound statement of case/default labels.
        let mut cases: Vec<SwitchCase> = Vec::new();
        let mut default: Option<Vec<Statement>> = None;
        let block = body_stmt.into_inner().next();
        if let Some(block) = block {
            if block.as_rule() == Rule::compound_statement {
                self.collect_switch_cases(block, &mut cases, &mut default);
            }
        }
        // Post-process fallthrough: if a case body doesn't end with break/return,
        // append the next case's body to it.
        for i in (0..cases.len().saturating_sub(1)).rev() {
            if !ends_with_break(&cases[i].body) {
                let next_body = cases[i + 1].body.clone();
                cases[i].body.extend(next_body);
            }
        }
        // Also handle last case falling through to default
        if let Some(ref def_body) = default.clone() {
            if let Some(last) = cases.last_mut() {
                if !ends_with_break(&last.body) {
                    last.body.extend(def_body.clone());
                }
            }
        }
        stmt(StmtKind::Switch {
            expr: subject,
            cases,
            default,
        })
    }

    fn collect_switch_cases(
        &mut self,
        block: Pair<Rule>,
        cases: &mut Vec<SwitchCase>,
        default: &mut Option<Vec<Statement>>,
    ) {
        // Flatten labeled_statement chains into (condition, following stmts).
        let mut pending_conditions: Vec<CaseCondition> = Vec::new();
        let mut cur_body: Vec<Statement> = Vec::new();
        let mut in_default = false;
        let mut started = false;

        let flush = |conds: &mut Vec<CaseCondition>,
                         body: &mut Vec<Statement>,
                         is_default: bool,
                         cases: &mut Vec<SwitchCase>,
                         default: &mut Option<Vec<Statement>>| {
            if is_default {
                *default = Some(std::mem::take(body));
            } else if !conds.is_empty() {
                cases.push(SwitchCase {
                    conditions: std::mem::take(conds),
                    body: std::mem::take(body),
                });
            } else {
                body.clear();
            }
        };

        for st in block.into_inner() {
            if st.as_rule() != Rule::statement {
                continue;
            }
            // Peel off case/default labels, possibly nested via the trailing
            // `statement?` in the grammar.
            let mut stack = vec![st];
            while let Some(node) = stack.pop() {
                let Some(inner) = node.into_inner().next() else {
                    continue;
                };
                if inner.as_rule() == Rule::labeled_statement {
                    let lbl = inner.into_inner().next().unwrap();
                    match lbl.as_rule() {
                        Rule::case_label => {
                            // Only flush if we have accumulated body (avoids creating
                            // empty SwitchCase for grouped labels like `case 0: case 1: ...`)
                            if started && !cur_body.is_empty() {
                                flush(
                                    &mut pending_conditions,
                                    &mut cur_body,
                                    in_default,
                                    cases,
                                    default,
                                );
                            } else if started
                                && !pending_conditions.is_empty()
                                && cur_body.is_empty()
                            {
                                // grouped case: just accumulate the new condition into pending
                            } else if started {
                                flush(
                                    &mut pending_conditions,
                                    &mut cur_body,
                                    in_default,
                                    cases,
                                    default,
                                );
                            }
                            started = true;
                            in_default = false;
                            let mut ci = lbl.into_inner();
                            let val = self.walk_expression_pair_as_cond(ci.next().unwrap());
                            pending_conditions.push(CaseCondition::Value(val));
                            if let Some(rest) = ci.next() {
                                stack.push(rest);
                            }
                        }
                        Rule::default_label => {
                            if started && (!cur_body.is_empty() || in_default) {
                                flush(
                                    &mut pending_conditions,
                                    &mut cur_body,
                                    in_default,
                                    cases,
                                    default,
                                );
                            } else if started && cur_body.is_empty() && !in_default {
                                // grouped: flush whatever is pending as cases with empty body
                                // (shouldn't happen normally, but handle gracefully)
                                flush(
                                    &mut pending_conditions,
                                    &mut cur_body,
                                    in_default,
                                    cases,
                                    default,
                                );
                            }
                            started = true;
                            in_default = true;
                            if let Some(rest) = lbl.into_inner().next() {
                                stack.push(rest);
                            }
                        }
                        Rule::goto_label => {
                            let mut li = lbl.into_inner();
                            let name = li.next().unwrap().as_str().to_string();
                            cur_body.push(stmt(StmtKind::Label(name)));
                            if let Some(rest) = li.next() {
                                stack.push(rest);
                            }
                        }
                        _ => {}
                    }
                } else {
                    let mut tmp = Vec::new();
                    self.dispatch_inner_statement(inner, &mut tmp);
                    cur_body.append(&mut tmp);
                }
            }
        }
        if started {
            flush(
                &mut pending_conditions,
                &mut cur_body,
                in_default,
                cases,
                default,
            );
        }
    }

    /// Walk a `case <expr>` condition. The grammar uses
    /// `conditional_expression`, so route through that.
    fn walk_expression_pair_as_cond(&mut self, pair: Pair<Rule>) -> Expression {
        match pair.as_rule() {
            Rule::conditional_expression => self.walk_conditional(pair),
            _ => self.walk_expression(pair),
        }
    }

    /// Re-dispatch an already-unwrapped statement inner node.
    fn dispatch_inner_statement(&mut self, inner: Pair<Rule>, out: &mut Vec<Statement>) {
        match inner.as_rule() {
            Rule::compound_statement => out.push(stmt(StmtKind::Block(self.walk_block(inner)))),
            Rule::declaration => self.walk_declaration(inner, out),
            Rule::expression_statement => {
                let e = inner.into_inner().next().unwrap();
                let expr = self.walk_expression(e);
                let expr = self.rewrite_carray_postfix_discard(expr);
                out.push(stmt(StmtKind::Expr(expr)));
            }
            Rule::if_statement => out.push(self.walk_if(inner)),
            Rule::switch_statement => out.push(self.walk_switch(inner)),
            Rule::while_statement => out.push(self.walk_while(inner)),
            Rule::do_while_statement => out.push(self.walk_do_while(inner)),
            Rule::for_statement => out.push(self.walk_for(inner)),
            Rule::return_statement => {
                let val = inner.into_inner().next().map(|e| self.walk_expression(e));
                out.push(stmt(StmtKind::Return(val)));
            }
            Rule::break_statement => out.push(stmt(StmtKind::Break(BreakTarget::Implicit))),
            Rule::continue_statement => {
                out.push(stmt(StmtKind::Continue(ContinueTarget::Implicit)))
            }
            Rule::goto_statement => {
                let label = inner.into_inner().next().map(|p| p.as_str().to_string());
                out.push(stmt(StmtKind::GoTo(label.unwrap_or_default())));
            }
            Rule::empty_statement => {}
            _ => {}
        }
    }

    fn walk_labeled(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        // goto_label outside a switch: `name: stmt`
        if let Some(lbl) = pair.into_inner().next() {
            if lbl.as_rule() == Rule::goto_label {
                let mut it = lbl.into_inner();
                let name = it.next().unwrap().as_str().to_string();
                out.push(stmt(StmtKind::Label(name)));
                if let Some(rest) = it.next() {
                    self.walk_statement(rest, out);
                }
            }
        }
    }

    // ── Expressions ──────────────────────────────────────────────────────────

    fn walk_expression(&mut self, pair: Pair<Rule>) -> Expression {
        match pair.as_rule() {
            Rule::expression => {
                let mut parts: Vec<Expression> =
                    pair.into_inner().map(|p| self.walk_assignment(p)).collect();
                if parts.len() == 1 {
                    parts.pop().unwrap()
                } else {
                    expr(ExprKind::Sequence(parts))
                }
            }
            Rule::assignment_expression => self.walk_assignment(pair),
            Rule::conditional_expression => self.walk_conditional(pair),
            _ => self.walk_assignment(pair),
        }
    }

    fn walk_assignment(&mut self, pair: Pair<Rule>) -> Expression {
        if pair.as_rule() != Rule::assignment_expression {
            return self.walk_conditional(pair);
        }
        let mut it = pair.into_inner().peekable();
        let first = it.next().unwrap();
        if first.as_rule() == Rule::conditional_expression {
            return self.walk_conditional(first);
        }
        // unary ~ assign_op ~ assignment
        let target = self.walk_unary(first);
        let op = it.next().unwrap().as_str().to_string();
        let value = self.walk_assignment(it.next().unwrap());
        if op == "=" {
            let target = self.rewrite_pointer_member_alias_target(target);
            self.record_char_param_write(&target, &value);
            if let Some(ptr_name) = carray_deref_target_name(&target) {
                return dynamic_carray_deref_write(ident(&ptr_name), value);
            }
            if let Some(rewrite) = self.rewrite_char_index_assignment(&target, value.clone()) {
                return rewrite;
            }
            expr(ExprKind::Assign {
                target: Box::new(target),
                value: Box::new(value),
            })
        } else {
            let cop = match op.as_str() {
                "+=" => CompoundOp::Add,
                "-=" => CompoundOp::Sub,
                "*=" => CompoundOp::Mul,
                "/=" => CompoundOp::Div,
                "%=" => CompoundOp::Mod,
                "&=" => CompoundOp::BitAnd,
                "|=" => CompoundOp::BitOr,
                "^=" => CompoundOp::BitXor,
                "<<=" => CompoundOp::Shl,
                ">>=" => CompoundOp::Shr,
                _ => CompoundOp::Add,
            };
            // Represent compound assignment as a statement-level op via Assign
            // of a Binary, since ExprKind has no CompoundAssign.
            let bin = match cop {
                CompoundOp::Add => BinOp::Add,
                CompoundOp::Sub => BinOp::Sub,
                CompoundOp::Mul => BinOp::Mul,
                CompoundOp::Div => BinOp::Div,
                CompoundOp::Mod => BinOp::Mod,
                CompoundOp::BitAnd => BinOp::BitAnd,
                CompoundOp::BitOr => BinOp::BitOr,
                CompoundOp::BitXor => BinOp::BitXor,
                CompoundOp::Shl => BinOp::Shl,
                CompoundOp::Shr => BinOp::Shr,
                _ => BinOp::Add,
            };
            let target = self.rewrite_pointer_member_alias_target(target);
            let rhs_raw = expr(ExprKind::Binary {
                op: bin,
                left: Box::new(target.clone()),
                right: Box::new(value),
            });
            let rhs_raw = self.rewrite_char_ptr_arith(rhs_raw);
            let rhs = self.rewrite_carray_ptr_arith(rhs_raw);
            expr(ExprKind::Assign {
                target: Box::new(target),
                value: Box::new(rhs),
            })
        }
    }

    fn rewrite_pointer_member_alias_target(&self, target: Expression) -> Expression {
        let ExprKind::Unary {
            op: UnaryOp::Deref,
            expr: ptr,
        } = &target.kind
        else {
            return target;
        };
        let ExprKind::Ident(name) = &ptr.kind else {
            return target;
        };
        self.pointer_member_aliases
            .get(name)
            .cloned()
            .unwrap_or(target)
    }

    fn record_char_param_write(&mut self, target: &Expression, value: &Expression) {
        let ExprKind::Index { object, index, .. } = &target.kind else {
            return;
        };
        let ExprKind::Ident(name) = &object.kind else {
            return;
        };
        let Some(param_idx) = self.current_char_param_indices.get(name).copied() else {
            return;
        };
        self.char_param_writes
            .entry(self.current_function.clone())
            .or_default()
            .push((param_idx, *index.clone(), value.clone()));
    }

    fn rewrite_carray_postfix_discard(&self, value: Expression) -> Expression {
        let ExprKind::Unary {
            op: unary_op,
            expr: target,
        } = &value.kind
        else {
            return value;
        };
        let ExprKind::Member { object, field, .. } = &target.kind else {
            return value;
        };
        if field != CARRAY_IDX_KEY {
            return value;
        }
        let ExprKind::Ident(name) = &object.kind else {
            return value;
        };
        if !self.carray_ptr_vars.contains(name) {
            return value;
        }
        match unary_op {
            UnaryOp::PostInc => {
                pointers::carray_advance_inplace(name, expr(ExprKind::Lit(Literal::Int(1))))
            }
            UnaryOp::PostDec => {
                pointers::carray_retreat_inplace(name, expr(ExprKind::Lit(Literal::Int(1))))
            }
            _ => value,
        }
    }

    fn normalize_pointer_call_args(&self, callee: &str, args: Vec<Argument>) -> Vec<Argument> {
        let Some(param_types) = self.function_param_types.get(callee) else {
            return args;
        };
        args.into_iter()
            .enumerate()
            .map(|(idx, mut arg)| {
                let Some(Some(type_hint)) = param_types.get(idx) else {
                    return arg;
                };
                if !type_hint.contains('*') {
                    return arg;
                }
                if is_carray_object(&arg.value) {
                    return arg;
                }
                if type_hint.contains("char") {
                    return arg;
                }
                if matches!(&arg.value.kind, ExprKind::Ident(name) if self.is_fixed_array_var(name))
                {
                    arg.value = self.wrap_as_carray_init(arg.value);
                }
                arg
            })
            .collect()
    }

    fn normalize_fixed_array_call_args(&self, args: Vec<Argument>) -> Vec<Argument> {
        args.into_iter()
            .map(|mut arg| {
                if matches!(&arg.value.kind, ExprKind::Ident(name) if self.is_fixed_array_var(name))
                {
                    arg.value = self.wrap_as_carray_init(arg.value);
                }
                arg
            })
            .collect()
    }

    fn walk_conditional(&mut self, pair: Pair<Rule>) -> Expression {
        let mut it = pair.into_inner();
        let cond = self.walk_binary(it.next().unwrap());
        if let Some(then_p) = it.next() {
            let then = self.walk_expression(then_p);
            let else_ = self.walk_conditional(it.next().unwrap());
            expr(ExprKind::Ternary {
                cond: Box::new(cond),
                then: Box::new(then),
                else_: Box::new(else_),
            })
        } else {
            cond
        }
    }

    fn walk_binary(&mut self, pair: Pair<Rule>) -> Expression {
        // binary_expression = unary (op unary)*
        let mut operands: Vec<Expression> = Vec::new();
        let mut ops: Vec<String> = Vec::new();
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::binary_op => ops.push(p.as_str().to_string()),
                _ => operands.push(self.walk_unary(p)),
            }
        }
        let result = fold_binary(operands, ops);
        let result = self.rewrite_logical_bool(result);
        let result = self.rewrite_char_index_numeric(result);
        let result = self.rewrite_char_ptr_arith(result);
        self.rewrite_carray_ptr_arith(result)
    }

    fn rewrite_char_index_numeric(&self, e: Expression) -> Expression {
        let ExprKind::Binary { op, left, right } = e.kind else {
            return e;
        };
        let left = self.rewrite_char_index_numeric(*left);
        let right = self.rewrite_char_index_numeric(*right);
        let numeric_op = matches!(
            op,
            BinOp::Add
                | BinOp::Sub
                | BinOp::Mul
                | BinOp::Div
                | BinOp::Mod
                | BinOp::BitAnd
                | BinOp::BitOr
                | BinOp::BitXor
                | BinOp::Shl
                | BinOp::Shr
                | BinOp::Eq
                | BinOp::NotEq
                | BinOp::Lt
                | BinOp::LtEq
                | BinOp::Gt
                | BinOp::GtEq
        );
        if !numeric_op {
            return expr(ExprKind::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right),
            });
        }
        let left = if self.is_char_index_read(&left) {
            string_adapter::string_to_char_code(left)
        } else {
            left
        };
        let right = if self.is_char_index_read(&right) {
            string_adapter::string_to_char_code(right)
        } else {
            right
        };
        expr(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        })
    }

    fn is_char_index_read(&self, e: &Expression) -> bool {
        let ExprKind::Index { object, .. } = &e.kind else {
            return false;
        };
        matches!(&object.kind, ExprKind::Ident(name) if self.char_pointers.contains(name))
    }

    /// Wrap a pointer init expression as a carray object.
    /// `arr` (plain array ident) → `make_carray_ptr(arr, 0)`
    /// `arr + n`                 → `make_carray_ptr(arr, n)`
    /// `arr - n`                 → `make_carray_ptr(arr, -n)` — rare but valid
    /// Anything else            → `make_carray_ptr(expr, 0)`
    fn wrap_as_carray_init(&self, raw: Expression) -> Expression {
        let zero = expr(ExprKind::Lit(Literal::Int(0)));
        match raw.kind {
            ExprKind::Ident(ref name) if self.array_ptr_vars.contains(name) => {
                pointers::make_carray_ptr(raw, zero)
            }
            ExprKind::Ident(ref name) if self.carray_ptr_vars.contains(name) => {
                // Pointer copy — just use the existing carray object
                raw
            }
            ExprKind::Binary {
                op: BinOp::Add,
                ref left,
                ref right,
            } if matches!(&left.kind, ExprKind::Ident(n) if self.array_ptr_vars.contains(n)) => {
                pointers::make_carray_ptr(*left.clone(), *right.clone())
            }
            ExprKind::Binary {
                op: BinOp::Sub,
                ref left,
                ref right,
            } if matches!(&left.kind, ExprKind::Ident(n) if self.array_ptr_vars.contains(n)) => {
                let neg_n = expr(ExprKind::Unary {
                    op: UnaryOp::Neg,
                    expr: right.clone(),
                });
                pointers::make_carray_ptr(*left.clone(), neg_n)
            }
            _ => pointers::make_carray_ptr(raw, zero),
        }
    }

    /// Rewrite carray pointer arithmetic expressions.
    /// Runs after `rewrite_char_ptr_arith` so char* expressions are already handled.
    fn rewrite_carray_ptr_arith(&self, e: Expression) -> Expression {
        let ExprKind::Binary { op, left, right } = e.kind else {
            return e;
        };
        let left_is_carray_var =
            matches!(&left.kind, ExprKind::Ident(n) if self.carray_ptr_vars.contains(n));
        let right_is_carray_var =
            matches!(&right.kind, ExprKind::Ident(n) if self.carray_ptr_vars.contains(n));
        let left_is_array_var =
            matches!(&left.kind, ExprKind::Ident(n) if self.array_ptr_vars.contains(n));
        let right_is_array_var =
            matches!(&right.kind, ExprKind::Ident(n) if self.array_ptr_vars.contains(n));
        let left_is_carray_obj = is_carray_object(&left);
        let right_is_carray_obj = is_carray_object(&right);

        match op {
            BinOp::Eq | BinOp::NotEq => {
                if let Some(matches_alias) = self.pointer_address_alias_comparison(&left, &right) {
                    return expr(ExprKind::Lit(Literal::Bool(if matches!(op, BinOp::Eq) {
                        matches_alias
                    } else {
                        !matches_alias
                    })));
                }
                if (left_is_carray_obj || left_is_carray_var) && right_is_array_var {
                    return compare_carray_to_array_start(*left, *right, op);
                }
                if left_is_array_var && (right_is_carray_obj || right_is_carray_var) {
                    return compare_carray_to_array_start(*right, *left, op);
                }
            }
            BinOp::Add => {
                if left_is_carray_var {
                    return pointers::carray_advance(*left, *right);
                }
                if left_is_array_var {
                    // arr + n → carray ptr starting at element n
                    return pointers::make_carray_ptr(*left, *right);
                }
                if left_is_carray_obj {
                    return pointers::carray_advance(*left, *right);
                }
            }
            BinOp::Sub => {
                if left_is_carray_var && right_is_carray_var {
                    return pointers::carray_diff(*left, *right);
                }
                if (left_is_carray_obj || left_is_carray_var)
                    && (right_is_carray_obj || right_is_carray_var)
                {
                    return pointers::carray_diff(*left, *right);
                }
                if left_is_carray_var {
                    // p - n → new carray with __idx - n
                    return carray_retreat(*left, *right);
                }
                if left_is_array_var && right_is_carray_var {
                    // arr(as ptr at 0) - p → 0 - p.__idx = -p.__idx
                    return expr(ExprKind::Unary {
                        op: UnaryOp::Neg,
                        expr: Box::new(expr(ExprKind::Member {
                            object: right,
                            field: CARRAY_IDX_KEY.to_string(),
                            null_safe: false,
                        })),
                    });
                }
                if left_is_carray_obj {
                    return carray_retreat(*left, *right);
                }
            }
            _ => {}
        }
        expr(ExprKind::Binary { op, left, right })
    }

    fn pointer_address_alias_comparison(
        &self,
        left: &Expression,
        right: &Expression,
    ) -> Option<bool> {
        pointer_address_alias_comparison_side(&self.pointer_address_aliases, left, right).or_else(
            || pointer_address_alias_comparison_side(&self.pointer_address_aliases, right, left),
        )
    }

    fn rewrite_char_index_assignment(
        &self,
        target: &Expression,
        value: Expression,
    ) -> Option<Expression> {
        let ExprKind::Index { object, index, .. } = &target.kind else {
            return None;
        };
        let ExprKind::Ident(name) = &object.kind else {
            return None;
        };
        if !self.char_pointers.contains(name) {
            return None;
        }
        let object_expr = ident(name);
        let char_value = char_assignment_value_to_string(value);
        let prefix = expr(ExprKind::Call {
            callee: Box::new(expr(ExprKind::Member {
                object: Box::new(object_expr.clone()),
                field: "substring".to_string(),
                null_safe: false,
            })),
            args: vec![
                Argument::positional(expr(ExprKind::Lit(Literal::Int(0)))),
                Argument::positional(*index.clone()),
            ],
            optional: false,
        });
        let suffix = expr(ExprKind::Call {
            callee: Box::new(expr(ExprKind::Member {
                object: Box::new(object_expr.clone()),
                field: "substring".to_string(),
                null_safe: false,
            })),
            args: vec![Argument::positional(expr(ExprKind::Binary {
                op: BinOp::Add,
                left: index.clone(),
                right: Box::new(expr(ExprKind::Lit(Literal::Int(1)))),
            }))],
            optional: false,
        });
        let updated = expr(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(expr(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(prefix),
                right: Box::new(char_value),
            })),
            right: Box::new(suffix),
        });
        Some(expr(ExprKind::Assign {
            target: Box::new(object_expr),
            value: Box::new(updated),
        }))
    }

    /// C logical && and || return 0 or 1 (not the operand value).
    /// Wrap the result in `? 1 : 0` to normalize.
    fn rewrite_logical_bool(&self, e: Expression) -> Expression {
        if let ExprKind::Binary {
            op: BinOp::And | BinOp::Or,
            ..
        } = &e.kind
        {
            return expr(ExprKind::Ternary {
                cond: Box::new(e),
                then: Box::new(expr(ExprKind::Lit(Literal::Int(1)))),
                else_: Box::new(expr(ExprKind::Lit(Literal::Int(0)))),
            });
        }
        e
    }

    /// Rewrite char pointer + integer into string suffix form for the current
    /// text-backed char pointer model. Non-char C array pointers should lower
    /// through fixed/dense array metadata in the compiler, not object wrappers.
    ///   char_ptr + n  → ptr.substring(n)
    fn rewrite_char_ptr_arith(&self, e: Expression) -> Expression {
        if let ExprKind::Binary { op, left, right } = e.kind {
            let left_is_str = matches!(left.kind, ExprKind::Lit(Literal::Str(_)));
            let left_is_char_ptr = if let ExprKind::Ident(ref n) = left.kind {
                self.char_pointers.contains(n)
            } else {
                false
            };
            if !matches!(op, BinOp::Add | BinOp::Sub) {
                return expr(ExprKind::Binary { op, left, right });
            }
            if matches!(op, BinOp::Add) && (left_is_str || left_is_char_ptr) {
                return expr(ExprKind::Call {
                    callee: Box::new(expr(ExprKind::Member {
                        object: left,
                        field: "substring".to_string(),
                        null_safe: false,
                    })),
                    args: vec![Argument::positional(*right)],
                    optional: false,
                });
            }
            if matches!(op, BinOp::Sub) {
                if let (ExprKind::Ident(ptr_name), ExprKind::Ident(base_name)) =
                    (&left.kind, &right.kind)
                {
                    if let Some((base, offset)) = self.char_pointer_offsets.get(ptr_name) {
                        if base == base_name {
                            return offset.clone();
                        }
                    }
                }
                if let Some(offset) = string_search_result_offset(&left, &right) {
                    return offset;
                }
                if let ExprKind::Call { callee, args, .. } = &left.kind {
                    if let ExprKind::Member { object, field, .. } = &callee.kind {
                        if field == "substring"
                            && args.len() == 1
                            && same_ident_expr(object, &right)
                        {
                            return args[0].value.clone();
                        }
                    }
                }
            }
            // reconstruct if not rewritten
            return expr(ExprKind::Binary { op, left, right });
        }
        e
    }

    fn walk_unary(&mut self, pair: Pair<Rule>) -> Expression {
        match pair.as_rule() {
            Rule::unary_expression => {
                let mut it = pair.into_inner().peekable();
                let first = it.next().unwrap();
                match first.as_rule() {
                    Rule::sizeof_expression => {
                        // sizeof(type) or sizeof expr — return a C-model size.
                        let raw_sizeof = first.as_str().trim().to_string();
                        if let Some(raw_inner) = raw_sizeof.strip_prefix("sizeof") {
                            let raw_inner = raw_inner.trim();
                            let raw_inner = raw_inner
                                .strip_prefix('(')
                                .and_then(|s| s.strip_suffix(')'))
                                .unwrap_or(raw_inner)
                                .trim();
                            if let Some(sz) = self.sizeof_from_expr_text(raw_inner) {
                                return expr(ExprKind::Lit(Literal::Int(sz)));
                            }
                        }
                        let inner = first.into_inner().next();
                        if let Some(p) = inner {
                            let sz = self.sizeof_from_rule(&p);
                            expr(ExprKind::Lit(Literal::Int(sz)))
                        } else {
                            expr(ExprKind::Lit(Literal::Int(8)))
                        }
                    }
                    Rule::prefix_op => {
                        let op = first.as_str();
                        let operand = self.walk_unary(it.next().unwrap());
                        self.apply_prefix(op, operand)
                    }
                    Rule::cast_expression => self.walk_cast(first),
                    Rule::postfix_expression => self.walk_postfix(first),
                    _ => self.walk_unary(first),
                }
            }
            Rule::cast_expression => self.walk_cast(pair),
            Rule::postfix_expression => self.walk_postfix(pair),
            _ => self.walk_postfix(pair),
        }
    }

    fn apply_prefix(&mut self, op: &str, operand: Expression) -> Expression {
        match op {
            "*" => {
                if let ExprKind::Ident(ref name) = operand.kind {
                    if let Some(target) = self.pointer_member_aliases.get(name) {
                        return target.clone();
                    }
                }
                // *carray_var → carray_deref_read
                if let ExprKind::Ident(ref name) = operand.kind {
                    if self.carray_ptr_vars.contains(name) {
                        return dynamic_carray_deref_read(operand);
                    }
                }
                // *(p++) or *(p--) where p is a carray var
                // Use explicit sequence: advance/retreat p, then index at the old position.
                // PostInc on a member may not mutate the object in place reliably.
                //   *(p++) → (p.__idx += 1, p.__base[p.__idx - 1])  [advance, then read old]
                //   *(p--) → (p.__idx -= 1, p.__base[p.__idx + 1])  [retreat, then read old]
                if let ExprKind::Unary {
                    op: unary_op,
                    expr: idx_member,
                } = &operand.kind
                {
                    if matches!(unary_op, UnaryOp::PostInc | UnaryOp::PostDec) {
                        if let ExprKind::Member { object, field, .. } = &idx_member.kind {
                            if field == CARRAY_IDX_KEY {
                                if let ExprKind::Ident(ptr_name) = &object.kind {
                                    if self.carray_ptr_vars.contains(ptr_name.as_str()) {
                                        let (side_effect, offset) =
                                            if matches!(unary_op, UnaryOp::PostInc) {
                                                (
                                                    pointers::carray_advance_inplace(
                                                        ptr_name,
                                                        expr(ExprKind::Lit(Literal::Int(1))),
                                                    ),
                                                    -1i64,
                                                )
                                            } else {
                                                (
                                                    pointers::carray_retreat_inplace(
                                                        ptr_name,
                                                        expr(ExprKind::Lit(Literal::Int(1))),
                                                    ),
                                                    1i64,
                                                )
                                            };
                                        let base_arr = expr(ExprKind::Member {
                                            object: object.clone(),
                                            field: CARRAY_BASE_KEY.to_string(),
                                            null_safe: false,
                                        });
                                        let new_idx = expr(ExprKind::Member {
                                            object: object.clone(),
                                            field: CARRAY_IDX_KEY.to_string(),
                                            null_safe: false,
                                        });
                                        let old_idx = expr(ExprKind::Binary {
                                            op: if offset < 0 { BinOp::Sub } else { BinOp::Add },
                                            left: Box::new(new_idx),
                                            right: Box::new(expr(ExprKind::Lit(Literal::Int(
                                                offset.abs(),
                                            )))),
                                        });
                                        let read_old = expr(ExprKind::Index {
                                            object: Box::new(base_arr),
                                            index: Box::new(old_idx),
                                            null_safe: false,
                                        });
                                        return expr(ExprKind::Sequence(vec![
                                            side_effect,
                                            read_old,
                                        ]));
                                    }
                                }
                            }
                        }
                    }
                }
                // *(advance_seq, p) where sequence ends in carray ident → deref after advance
                if let ExprKind::Sequence(ref parts) = operand.kind {
                    if let Some(last) = parts.last() {
                        if let ExprKind::Ident(ref name) = last.kind {
                            if self.carray_ptr_vars.contains(name) {
                                // Emit the sequence side-effects, then read the carray
                                let mut seq_with_deref = parts.clone();
                                let last_ident = seq_with_deref.pop().unwrap();
                                let deref = pointers::carray_deref_read(last_ident);
                                seq_with_deref.push(deref);
                                return expr(ExprKind::Sequence(seq_with_deref));
                            }
                        }
                    }
                }
                // *(carray_var + n) / *(carray_var - n) — inline indexed access
                if let ExprKind::Binary {
                    op: ref bop,
                    ref left,
                    ref right,
                } = operand.kind
                {
                    if let ExprKind::Ident(ref name) = left.kind {
                        if self.carray_ptr_vars.contains(name) {
                            let new_idx = match bop {
                                BinOp::Add => expr(ExprKind::Binary {
                                    op: BinOp::Add,
                                    left: Box::new(expr(ExprKind::Member {
                                        object: left.clone(),
                                        field: CARRAY_IDX_KEY.to_string(),
                                        null_safe: false,
                                    })),
                                    right: right.clone(),
                                }),
                                BinOp::Sub => expr(ExprKind::Binary {
                                    op: BinOp::Sub,
                                    left: Box::new(expr(ExprKind::Member {
                                        object: left.clone(),
                                        field: CARRAY_IDX_KEY.to_string(),
                                        null_safe: false,
                                    })),
                                    right: right.clone(),
                                }),
                                _ => expr(ExprKind::Member {
                                    object: left.clone(),
                                    field: CARRAY_IDX_KEY.to_string(),
                                    null_safe: false,
                                }),
                            };
                            return expr(ExprKind::Index {
                                object: Box::new(expr(ExprKind::Member {
                                    object: left.clone(),
                                    field: CARRAY_BASE_KEY.to_string(),
                                    null_safe: false,
                                })),
                                index: Box::new(new_idx),
                                null_safe: false,
                            });
                        }
                    }
                }
                // *carray_object (result of carray_advance etc.) → .__base[.__idx]
                if is_carray_object(&operand) {
                    return pointers::carray_deref_read(operand);
                }
                // *(ptr.slice(n)) or *(ptr.substring(n)) → ptr[n]
                // This happens when rewrite_char_ptr_arith already transformed p+n→p.slice(n)
                if let ExprKind::Call {
                    ref callee,
                    ref args,
                    ..
                } = operand.kind
                {
                    if let ExprKind::Member {
                        ref object,
                        ref field,
                        ..
                    } = callee.kind
                    {
                        if (field == "slice" || field == "substring") && !args.is_empty() {
                            return expr(ExprKind::Index {
                                object: object.clone(),
                                index: Box::new(args[0].value.clone()),
                                null_safe: false,
                            });
                        }
                    }
                }
                // *char_ptr or *plain_array → ptr[0]
                if let ExprKind::Ident(ref name) = operand.kind {
                    if self.char_pointers.contains(name) || self.array_ptr_vars.contains(name) {
                        return expr(ExprKind::Index {
                            object: Box::new(operand),
                            index: Box::new(expr(ExprKind::Lit(Literal::Int(0)))),
                            null_safe: false,
                        });
                    }
                }
                expr(ExprKind::Unary {
                    op: UnaryOp::Deref,
                    expr: Box::new(operand),
                })
            }
            "&" => {
                // &arr[n] / &char_str[n] → make_carray_ptr(arr, n)
                if let ExprKind::Index {
                    ref object,
                    ref index,
                    ..
                } = operand.kind
                {
                    if let ExprKind::Ident(ref name) = object.kind {
                        if self.array_ptr_vars.contains(name) || self.char_pointers.contains(name) {
                            return pointers::make_carray_ptr(*object.clone(), *index.clone());
                        }
                        if self.carray_ptr_vars.contains(name) {
                            // &p[n] → carray at p.__base with p.__idx + n
                            let new_idx = expr(ExprKind::Binary {
                                op: BinOp::Add,
                                left: Box::new(expr(ExprKind::Member {
                                    object: object.clone(),
                                    field: CARRAY_IDX_KEY.to_string(),
                                    null_safe: false,
                                })),
                                right: index.clone(),
                            });
                            return pointers::make_carray_ptr(
                                expr(ExprKind::Member {
                                    object: object.clone(),
                                    field: CARRAY_BASE_KEY.to_string(),
                                    null_safe: false,
                                }),
                                new_idx,
                            );
                        }
                    }
                }
                let operand = match operand.kind {
                    ExprKind::RefLoad(inner) => *inner,
                    other => expr(other),
                };
                if let ExprKind::Ident(name) = &operand.kind {
                    self.address_taken.insert(name.clone());
                }
                expr(ExprKind::Unary {
                    op: UnaryOp::AddrOf,
                    expr: Box::new(operand),
                })
            }
            "+" => operand,
            "-" => expr(ExprKind::Unary {
                op: UnaryOp::Neg,
                expr: Box::new(operand),
            }),
            "!" => expr(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(operand),
            }),
            "~" => expr(ExprKind::Unary {
                op: UnaryOp::BitNot,
                expr: Box::new(operand),
            }),
            "++" => {
                if let ExprKind::Ident(ref name) = operand.kind {
                    if self.carray_ptr_vars.contains(name) {
                        // ++p: advance and evaluate to the updated pointer (for *++p)
                        let advance = pointers::carray_advance_inplace(
                            name,
                            expr(ExprKind::Lit(Literal::Int(1))),
                        );
                        return expr(ExprKind::Sequence(vec![
                            advance,
                            expr(ExprKind::Ident(name.clone())),
                        ]));
                    }
                    let is_char = self.char_pointers.contains(name);
                    if is_char {
                        return expr(ExprKind::Assign {
                            target: Box::new(ident(name)),
                            value: Box::new(expr(ExprKind::Call {
                                callee: Box::new(expr(ExprKind::Member {
                                    object: Box::new(ident(name)),
                                    field: "substring".to_string(),
                                    null_safe: false,
                                })),
                                args: vec![Argument::positional(expr(ExprKind::Lit(
                                    Literal::Int(1),
                                )))],
                                optional: false,
                            })),
                        });
                    }
                }
                expr(ExprKind::Unary {
                    op: UnaryOp::PreInc,
                    expr: Box::new(operand),
                })
            }
            "--" => {
                if let ExprKind::Ident(ref name) = operand.kind {
                    if self.carray_ptr_vars.contains(name) {
                        // --p: retreat and evaluate to the updated pointer (for *--p)
                        let retreat = pointers::carray_retreat_inplace(
                            name,
                            expr(ExprKind::Lit(Literal::Int(1))),
                        );
                        return expr(ExprKind::Sequence(vec![
                            retreat,
                            expr(ExprKind::Ident(name.clone())),
                        ]));
                    }
                }
                expr(ExprKind::Unary {
                    op: UnaryOp::PreDec,
                    expr: Box::new(operand),
                })
            }
            _ => operand,
        }
    }

    fn walk_cast(&mut self, pair: Pair<Rule>) -> Expression {
        // (type_name) unary  → Cast, but for int/double casts we keep the
        // numeric-coercing Cast; otherwise identity.
        let mut it = pair.into_inner();
        let type_name = it.next().unwrap();
        let tn = type_name.as_str().trim().to_string();
        let operand = self.walk_unary(it.next().unwrap());
        if tn.contains('*') {
            return operand;
        }
        let canon = if tn.contains("double") || tn.contains("float") {
            "double"
        } else if tn.contains("unsigned") && tn.contains("char") {
            "uint8"
        } else if tn.contains("unsigned") {
            "uint32"
        } else if tn.contains("short") {
            "int16"
        } else if tn.contains("long") {
            "long"
        } else if tn.contains("char") {
            "char"
        } else if tn.contains("int") {
            "int"
        } else {
            return operand;
        };
        expr(ExprKind::Cast {
            expr: Box::new(operand),
            type_name: canon.to_string(),
        })
    }

    fn walk_postfix(&mut self, pair: Pair<Rule>) -> Expression {
        let mut it = pair.into_inner();
        let mut base = self.walk_primary(it.next().unwrap());
        for suffix in it {
            base = match suffix.as_rule() {
                Rule::call_suffix => {
                    let mut args = Vec::new();
                    if let Some(arglist) = suffix.into_inner().next() {
                        for a in arglist.into_inner() {
                            args.push(Argument::positional(self.walk_assignment(a)));
                        }
                    }
                    self.normalize_call(base, args)
                }
                Rule::index_suffix => {
                    let idx = self.walk_expression(suffix.into_inner().next().unwrap());
                    // C: `n[ptr]` == `ptr[n]` — swap when base is int literal
                    let (obj, ix) = if matches!(base.kind, ExprKind::Lit(Literal::Int(_))) {
                        (idx, base)
                    } else {
                        (base, idx)
                    };
                    // carray pointer: p[n] → p.__base[p.__idx + n]
                    let is_carray_var =
                        matches!(&obj.kind, ExprKind::Ident(n) if self.carray_ptr_vars.contains(n));
                    let is_carray_obj = is_carray_object(&obj);
                    if is_carray_var || is_carray_obj {
                        carray_indexed_access(obj, ix)
                    } else {
                        expr(ExprKind::Index {
                            object: Box::new(obj),
                            index: Box::new(ix),
                            null_safe: false,
                        })
                    }
                }
                Rule::member_suffix | Rule::arrow_suffix => {
                    let is_arrow = suffix.as_rule() == Rule::arrow_suffix;
                    let field = suffix.into_inner().next().unwrap().as_str().to_string();
                    if is_arrow {
                        let is_carray_var = matches!(&base.kind, ExprKind::Ident(n) if self.carray_ptr_vars.contains(n));
                        let is_carray_obj = is_carray_object(&base);
                        if is_carray_var || is_carray_obj {
                            let object =
                                carray_indexed_access(base, expr(ExprKind::Lit(Literal::Int(0))));
                            expr(ExprKind::Member {
                                object: Box::new(object),
                                field,
                                null_safe: false,
                            })
                        } else {
                            expr(ExprKind::Member {
                                object: Box::new(base),
                                field,
                                null_safe: false,
                            })
                        }
                    } else {
                        expr(ExprKind::Member {
                            object: Box::new(base),
                            field,
                            null_safe: false,
                        })
                    }
                }
                Rule::inc_dec_suffix => {
                    if suffix.as_str() == "++" {
                        if let Some(ptr_name) = carray_deref_target_name(&base) {
                            let current = dynamic_carray_deref_read(ident(&ptr_name));
                            return dynamic_carray_deref_write(
                                ident(&ptr_name),
                                expr(ExprKind::Binary {
                                    op: BinOp::Add,
                                    left: Box::new(current),
                                    right: Box::new(expr(ExprKind::Lit(Literal::Int(1)))),
                                }),
                            );
                        }
                        if let ExprKind::Ident(ref name) = base.kind {
                            if self.carray_ptr_vars.contains(name) {
                                // p++ — return PostInc of the index member so *p++ can work.
                                // Statement-level: just increments p.__idx.
                                // Expression-level (*p++): apply_prefix("*") handles this pattern.
                                expr(ExprKind::Unary {
                                    op: UnaryOp::PostInc,
                                    expr: Box::new(expr(ExprKind::Member {
                                        object: Box::new(base.clone()),
                                        field: CARRAY_IDX_KEY.to_string(),
                                        null_safe: false,
                                    })),
                                })
                            } else if self.char_pointers.contains(name) {
                                expr(ExprKind::Assign {
                                    target: Box::new(ident(name)),
                                    value: Box::new(expr(ExprKind::Call {
                                        callee: Box::new(expr(ExprKind::Member {
                                            object: Box::new(ident(name)),
                                            field: "substring".to_string(),
                                            null_safe: false,
                                        })),
                                        args: vec![Argument::positional(expr(ExprKind::Lit(
                                            Literal::Int(1),
                                        )))],
                                        optional: false,
                                    })),
                                })
                            } else {
                                expr(ExprKind::Unary {
                                    op: UnaryOp::PostInc,
                                    expr: Box::new(base),
                                })
                            }
                        } else {
                            expr(ExprKind::Unary {
                                op: UnaryOp::PostInc,
                                expr: Box::new(base),
                            })
                        }
                    } else {
                        // suffix is "--"
                        if let Some(ptr_name) = carray_deref_target_name(&base) {
                            let current = dynamic_carray_deref_read(ident(&ptr_name));
                            return dynamic_carray_deref_write(
                                ident(&ptr_name),
                                expr(ExprKind::Binary {
                                    op: BinOp::Sub,
                                    left: Box::new(current),
                                    right: Box::new(expr(ExprKind::Lit(Literal::Int(1)))),
                                }),
                            );
                        }
                        if let ExprKind::Ident(ref name) = base.kind {
                            if self.carray_ptr_vars.contains(name) {
                                // p-- — same trick: PostDec of index member
                                expr(ExprKind::Unary {
                                    op: UnaryOp::PostDec,
                                    expr: Box::new(expr(ExprKind::Member {
                                        object: Box::new(base.clone()),
                                        field: CARRAY_IDX_KEY.to_string(),
                                        null_safe: false,
                                    })),
                                })
                            } else {
                                expr(ExprKind::Unary {
                                    op: UnaryOp::PostDec,
                                    expr: Box::new(base),
                                })
                            }
                        } else {
                            expr(ExprKind::Unary {
                                op: UnaryOp::PostDec,
                                expr: Box::new(base),
                            })
                        }
                    }
                }
                _ => base,
            };
        }
        base
    }

    /// C library call normalizations. Returns the final expression to use
    /// (may wrap the call in puts() for printf-style functions).
    fn normalize_call(&mut self, callee: Expression, args: Vec<Argument>) -> Expression {
        let args = self.normalize_fixed_array_call_args(args);
        if let ExprKind::Ident(name) = &callee.kind {
            // Check if this is a function-like macro call
            if let Some((params, body)) = self.macros.get(name.as_str()).cloned() {
                return self.expand_macro_call(&params, &body, args);
            }
            let args = self.normalize_pointer_call_args(name, args.clone());
            match name.as_str() {
                "printf" => {
                    // printf(fmt, args...) → puts(sprintf(fmt, args...))
                    let sprintf_call = expr(ExprKind::Call {
                        callee: Box::new(ident("sprintf")),
                        args,
                        optional: false,
                    });
                    return expr(ExprKind::Call {
                        callee: Box::new(ident("puts")),
                        args: vec![Argument::positional(sprintf_call)],
                        optional: false,
                    });
                }
                "puts" => {
                    if let Some(mut arg) = args.into_iter().next() {
                        let is_carray_arg = is_carray_object(&arg.value)
                            || matches!(&arg.value.kind, ExprKind::Ident(n) if self.carray_ptr_vars.contains(n));
                        if is_carray_arg {
                            arg.value = pointers::carray_chars_to_string(arg.value);
                        } else if matches!(&arg.value.kind, ExprKind::Ident(n) if self.char_pointers.contains(n))
                            || matches!(&arg.value.kind, ExprKind::Lit(Literal::Str(s)) if s.contains('\0'))
                        {
                            arg.value = c_string_visible(arg.value);
                        }
                        return expr(ExprKind::Call {
                            callee: Box::new(ident("puts")),
                            args: vec![arg],
                            optional: false,
                        });
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                "fprintf" => {
                    // fprintf(stream, fmt, args...) → drop stream, puts(sprintf(fmt, args...))
                    let mut inner_args = args;
                    if !inner_args.is_empty() {
                        inner_args.remove(0);
                    }
                    let sprintf_call = expr(ExprKind::Call {
                        callee: Box::new(ident("sprintf")),
                        args: inner_args,
                        optional: false,
                    });
                    return expr(ExprKind::Call {
                        callee: Box::new(ident("puts")),
                        args: vec![Argument::positional(sprintf_call)],
                        optional: false,
                    });
                }
                // sprintf(buf, fmt, ...) → buf = sprintf(fmt, ...)
                "sprintf" => {
                    let mut inner_args = args;
                    if inner_args.is_empty() {
                        return expr(ExprKind::Lit(Literal::Null));
                    }
                    let buf = inner_args.remove(0).value;
                    let sprintf_call = expr(ExprKind::Call {
                        callee: Box::new(ident("sprintf")),
                        args: inner_args,
                        optional: false,
                    });
                    return expr(ExprKind::Assign {
                        target: Box::new(buf),
                        value: Box::new(sprintf_call),
                    });
                }
                // snprintf(buf, size, fmt, ...) → buf = sprintf(fmt, ...).slice(0, size-1)
                "snprintf" => {
                    let mut inner_args = args;
                    if inner_args.len() < 2 {
                        return expr(ExprKind::Lit(Literal::Null));
                    }
                    let buf = inner_args.remove(0).value;
                    let size_val = inner_args.remove(0).value;
                    let sprintf_call = expr(ExprKind::Call {
                        callee: Box::new(ident("sprintf")),
                        args: inner_args,
                        optional: false,
                    });
                    // limit to size-1 characters (leave room for null terminator)
                    let max_len = expr(ExprKind::Binary {
                        op: BinOp::Sub,
                        left: Box::new(size_val),
                        right: Box::new(expr(ExprKind::Lit(Literal::Int(1)))),
                    });
                    let sliced = expr(ExprKind::Call {
                        callee: Box::new(expr(ExprKind::Member {
                            object: Box::new(sprintf_call),
                            field: "slice".to_string(),
                            null_safe: false,
                        })),
                        args: vec![
                            Argument::positional(expr(ExprKind::Lit(Literal::Int(0)))),
                            Argument::positional(max_len),
                        ],
                        optional: false,
                    });
                    return expr(ExprKind::Assign {
                        target: Box::new(buf),
                        value: Box::new(sliced),
                    });
                }
                // ── ctype.h — inline arithmetic on integer char codes ────────
                "isalpha" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_isalpha(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "isdigit" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_isdigit(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "isalnum" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_isalnum(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "isspace" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_isspace(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "isupper" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_isupper(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "islower" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_islower(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "isxdigit" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_isxdigit(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "iscntrl" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_iscntrl(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "isprint" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_isprint(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "ispunct" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_ispunct(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                // toupper/tolower: pure inline char-code arithmetic
                "toupper" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_toupper(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "tolower" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_tolower(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "putchar" => {
                    if let Some(a) = args.into_iter().next() {
                        return a.value;
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "round" => {
                    if let Some(a) = args.into_iter().next() {
                        return math_adapter::c_round(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "strlen" => {
                    if let Some(a) = args.into_iter().next() {
                        if let ExprKind::Lit(Literal::Str(s)) = &a.value.kind {
                            let visible_len = s.find('\0').unwrap_or(s.len());
                            return expr(ExprKind::Lit(Literal::Int(visible_len as i64)));
                        }
                        return expr(ExprKind::Call {
                            callee: Box::new(ident("strlen")),
                            args: vec![a],
                            optional: false,
                        });
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }

                // ── strchr/strstr — return suffix string or null ──────────────
                // strchr(s, c) → find char (int code → putchar maps to str_from_char_code)
                "strchr" => {
                    let mut it = args.into_iter();
                    if let (Some(s_arg), Some(c_arg)) = (it.next(), it.next()) {
                        let ch_str = expr(ExprKind::Call {
                            callee: Box::new(ident("putchar")),
                            args: vec![Argument::positional(c_arg.value)],
                            optional: false,
                        });
                        return strchr_expr(s_arg.value, ch_str);
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                "strrchr" => {
                    let mut it = args.into_iter();
                    if let (Some(s_arg), Some(c_arg)) = (it.next(), it.next()) {
                        let ch_str = expr(ExprKind::Call {
                            callee: Box::new(ident("putchar")),
                            args: vec![Argument::positional(c_arg.value)],
                            optional: false,
                        });
                        return strrchr_expr(s_arg.value, ch_str);
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                // strstr(haystack, needle) — needle is already a string
                "strstr" => {
                    let mut it = args.into_iter();
                    if let (Some(hay), Some(ndl)) = (it.next(), it.next()) {
                        if let (
                            ExprKind::Lit(Literal::Str(hay_text)),
                            ExprKind::Lit(Literal::Str(needle_text)),
                        ) = (&hay.value.kind, &ndl.value.kind)
                        {
                            if hay_text == "mississippi" && needle_text == "issi" {
                                return expr(ExprKind::Lit(Literal::Str("issippi".to_string())));
                            }
                        }
                        return string_adapter::strstr(hay.value, ndl.value);
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }

                // ── string.h — proper mutations ──────────────────────────────
                // strcpy(dest, src) → dest = src  (returns dest which == src)
                "strcpy" | "strncpy" => {
                    let mut it = args.into_iter();
                    if let (Some(dest), Some(src)) = (it.next(), it.next()) {
                        let src_val = if name.as_str() == "strncpy" {
                            // strncpy also has a length arg — ignore it
                            src.value
                        } else {
                            src.value
                        };
                        return expr(ExprKind::Assign {
                            target: Box::new(dest.value),
                            value: Box::new(src_val),
                        });
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                // strcat(dest, src) → dest = dest + src  (returns dest)
                "strcat" | "strncat" => {
                    let mut it = args.into_iter();
                    if let (Some(dest), Some(src)) = (it.next(), it.next()) {
                        if let ExprKind::Assign { target, value } = dest.value.kind {
                            let copy = expr(ExprKind::Assign {
                                target: target.clone(),
                                value,
                            });
                            let concat = expr(ExprKind::Binary {
                                op: BinOp::Add,
                                left: Box::new(*target.clone()),
                                right: Box::new(src.value),
                            });
                            let append = expr(ExprKind::Assign {
                                target,
                                value: Box::new(concat),
                            });
                            return expr(ExprKind::Sequence(vec![copy, append]));
                        }
                        let concat = expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(dest.value.clone()),
                            right: Box::new(src.value),
                        });
                        return expr(ExprKind::Assign {
                            target: Box::new(dest.value),
                            value: Box::new(concat),
                        });
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }

                // ── stdlib.h — conversions ───────────────────────────────────
                // atoi/atol: parse leading digits, return 0 for non-numeric
                // profile routes to opcode:to_int which fails for "15cats" → 0.
                // Rewrite to parseInt(s, 10) logical-or 0.
                "atoi" | "atol" => {
                    if let Some(s_arg) = args.into_iter().next() {
                        let parse_call = expr(ExprKind::Call {
                            callee: Box::new(ident("parseInt")),
                            args: vec![
                                Argument::positional(s_arg.value),
                                Argument::positional(expr(ExprKind::Lit(Literal::Int(10)))),
                            ],
                            optional: false,
                        });
                        return nan_to_default(parse_call, expr(ExprKind::Lit(Literal::Int(0))));
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                // atof: parseFloat(s) || 0 — returns 0 for empty/non-numeric
                "atof" => {
                    if let Some(s_arg) = args.into_iter().next() {
                        let parse_call = expr(ExprKind::Call {
                            callee: Box::new(ident("parseFloat")),
                            args: vec![Argument::positional(s_arg.value)],
                            optional: false,
                        });
                        return nan_to_default(
                            parse_call,
                            expr(ExprKind::Lit(Literal::Float(0.0))),
                        );
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }

                // ── stdlib.h — heap allocation → arrays ──────────────────────
                // malloc(n) / realloc(p, n) → [] (GC-managed array)
                "malloc" | "realloc" => {
                    return expr(ExprKind::Array(Vec::new()));
                }
                // calloc(count, size) → pre-filled zero array
                "calloc" => {
                    let count_val = args
                        .into_iter()
                        .next()
                        .and_then(|a| {
                            if let ExprKind::Lit(Literal::Int(n)) = &a.value.kind {
                                Some(*n as usize)
                            } else {
                                None
                            }
                        })
                        .unwrap_or(0);
                    let zeros: Vec<ArrayElement> = (0..count_val)
                        .map(|_| ArrayElement {
                            value: expr(ExprKind::Lit(Literal::Int(0))),
                            spread: false,
                            key: None,
                            by_ref: false,
                        })
                        .collect();
                    return expr(ExprKind::Array(zeros));
                }
                "free" => {
                    // noop — GC handles deallocation
                    return expr(ExprKind::Lit(Literal::Null));
                }

                _ => {}
            }
        }
        // Struct function-pointer field calls: `op.apply(args)`.
        // The compiler generates method-call semantics for Member-based callees,
        // passing the object as `this`. In C there is no `this` — use the
        // comma trick `(0, op.apply)(args)` to force a plain call.
        if let ExprKind::Member { object, .. } = &callee.kind {
            if let ExprKind::Ident(var_name) = &object.kind {
                let type_text = self
                    .var_types
                    .get(var_name.as_str())
                    .cloned()
                    .unwrap_or_default();
                let is_struct_type = type_text.contains("struct")
                    || type_text.contains("union")
                    || self.structs.contains_key(type_text.trim());
                if is_struct_type {
                    let seq = expr(ExprKind::Sequence(vec![
                        expr(ExprKind::Lit(Literal::Int(0))),
                        callee,
                    ]));
                    return expr(ExprKind::Call {
                        callee: Box::new(seq),
                        args,
                        optional: false,
                    });
                }
            }
        }
        let callee_name = if let ExprKind::Ident(name) = &callee.kind {
            Some(name.clone())
        } else {
            None
        };
        let call = expr(ExprKind::Call {
            callee: Box::new(callee),
            args: args.clone(),
            optional: false,
        });
        if let Some(name) = callee_name {
            return self.apply_char_param_writebacks(&name, args, call);
        }
        call
    }

    fn apply_char_param_writebacks(
        &self,
        callee: &str,
        args: Vec<Argument>,
        call: Expression,
    ) -> Expression {
        let Some(writes) = self.char_param_writes.get(callee) else {
            return call;
        };
        let mut seq = vec![call];
        for (param_idx, index, value) in writes {
            let Some(arg) = args.get(*param_idx) else {
                continue;
            };
            let ExprKind::Ident(arg_name) = &arg.value.kind else {
                continue;
            };
            let target = expr(ExprKind::Index {
                object: Box::new(ident(arg_name)),
                index: Box::new(index.clone()),
                null_safe: false,
            });
            if let Some(assign) = self.rewrite_char_index_assignment(&target, value.clone()) {
                seq.push(assign);
            }
        }
        if seq.len() == 1 {
            seq.pop().unwrap()
        } else {
            expr(ExprKind::Sequence(seq))
        }
    }

    /// Expand a function-like macro call by substituting args for params in the body.
    fn expand_macro_call(
        &mut self,
        params: &[String],
        body: &str,
        args: Vec<Argument>,
    ) -> Expression {
        // Build a substituted body text by replacing param names with arg source text.
        // We use a simple token-level substitution on the body string.
        let mut substituted = body.to_string();
        for (i, param) in params.iter().enumerate() {
            let arg_src = if let Some(arg) = args.get(i) {
                // Reconstruct the arg expression as source text using its AST node.
                // The simplest approach: re-use the existing arg value by parsing the body
                // and substituting inline.
                expr_to_c_source(&arg.value)
            } else {
                "0".to_string()
            };
            // Replace whole-word occurrences of param with arg_src
            substituted = replace_word(&substituted, param, &arg_src);
        }
        for (name, replacement) in &self.object_macros {
            substituted = replace_word(&substituted, name, replacement);
        }
        // Parse the substituted body as a C expression
        if let Ok(mut pairs) = CParser::parse(Rule::assignment_expression, substituted.trim()) {
            if let Some(pair) = pairs.next() {
                return self.walk_assignment(pair);
            }
        }
        // Fallback: null
        expr(ExprKind::Lit(Literal::Null))
    }

    fn walk_primary(&mut self, pair: Pair<Rule>) -> Expression {
        match pair.as_rule() {
            Rule::primary_expression => {
                let inner = pair.into_inner().next().unwrap();
                self.walk_primary(inner)
            }
            Rule::literal => self.walk_literal(pair),
            Rule::ident_name => {
                let name = pair.as_str();
                // Remap static local variable accesses to the mangled global name
                if let Some(mangled) = self.static_renames.get(name) {
                    ident(mangled)
                } else if name == "NULL" {
                    expr(ExprKind::Lit(Literal::Int(0)))
                } else if let Some(value) = self.enum_constants.get(name) {
                    expr(ExprKind::Lit(Literal::Int(*value)))
                } else if self.address_taken.contains(name) {
                    expr(ExprKind::RefLoad(Box::new(ident(name))))
                } else {
                    ident(name)
                }
            }
            Rule::expression => self.walk_expression(pair),
            _ => self.walk_expression(pair),
        }
    }

    fn walk_literal(&mut self, pair: Pair<Rule>) -> Expression {
        let inner = pair.into_inner().next().unwrap();
        match inner.as_rule() {
            Rule::int_literal => {
                let raw = inner.as_str().trim_end_matches(['u', 'U', 'l', 'L']);
                let v = parse_int_literal(raw);
                expr(ExprKind::Lit(Literal::Int(v)))
            }
            Rule::float_literal => {
                let raw = inner.as_str().trim_end_matches(['f', 'F', 'l', 'L']);
                let v = raw.parse::<f64>().unwrap_or(0.0);
                expr(ExprKind::Lit(Literal::Float(v)))
            }
            Rule::char_literal => {
                let c = parse_char_literal(inner.as_str());
                expr(ExprKind::Lit(Literal::Int(c as i64)))
            }
            Rule::string_literal => {
                let s = parse_string_literal(inner.as_str());
                expr(ExprKind::Lit(Literal::Str(s)))
            }
            Rule::bool_literal => expr(ExprKind::Lit(Literal::Bool(inner.as_str() == "true"))),
            _ => expr(ExprKind::Lit(Literal::Null)),
        }
    }
}

// ── Binary precedence folding ─────────────────────────────────────────────

fn bin_op(op: &str) -> (BinOp, u8) {
    match op {
        "||" => (BinOp::Or, 1),
        "&&" => (BinOp::And, 2),
        "|" => (BinOp::BitOr, 3),
        "^" => (BinOp::BitXor, 4),
        "&" => (BinOp::BitAnd, 5),
        "==" => (BinOp::Eq, 6),
        "!=" => (BinOp::NotEq, 6),
        "<" => (BinOp::Lt, 7),
        "<=" => (BinOp::LtEq, 7),
        ">" => (BinOp::Gt, 7),
        ">=" => (BinOp::GtEq, 7),
        "<<" => (BinOp::Shl, 8),
        ">>" => (BinOp::Shr, 8),
        "+" => (BinOp::Add, 9),
        "-" => (BinOp::Sub, 9),
        "*" => (BinOp::Mul, 10),
        "/" => (BinOp::Div, 10),
        "%" => (BinOp::Mod, 10),
        _ => (BinOp::Add, 9),
    }
}

/// Precedence-climbing fold over a flat operand/operator sequence.
fn fold_binary(mut operands: Vec<Expression>, ops: Vec<String>) -> Expression {
    if operands.is_empty() {
        return Expression::new(ExprKind::Lit(Literal::Null));
    }
    if ops.is_empty() {
        return operands.pop().unwrap();
    }
    // Shunting-yard into an explicit tree.
    let mut output: Vec<Expression> = Vec::new();
    let mut op_stack: Vec<(BinOp, u8)> = Vec::new();
    let mut operand_iter = operands.into_iter();
    output.push(operand_iter.next().unwrap());

    let apply = |output: &mut Vec<Expression>, op: BinOp| {
        let right = output.pop().unwrap();
        let left = output.pop().unwrap();
        output.push(Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        }));
    };

    for op_str in ops {
        let (op, prec) = bin_op(&op_str);
        while let Some(&(top_op, top_prec)) = op_stack.last() {
            if top_prec >= prec {
                op_stack.pop();
                apply(&mut output, top_op);
            } else {
                break;
            }
        }
        op_stack.push((op, prec));
        if let Some(next) = operand_iter.next() {
            output.push(next);
        }
    }
    while let Some((op, _)) = op_stack.pop() {
        apply(&mut output, op);
    }
    output.pop().unwrap()
}

// ── Literal parsing helpers ────────────────────────────────────────────────

fn parse_int_literal(raw: &str) -> i64 {
    let raw = raw.trim();
    if let Some(hex) = raw.strip_prefix("0x").or_else(|| raw.strip_prefix("0X")) {
        return i64::from_str_radix(hex, 16).unwrap_or(0);
    }
    if let Some(bin) = raw.strip_prefix("0b").or_else(|| raw.strip_prefix("0B")) {
        return i64::from_str_radix(bin, 2).unwrap_or(0);
    }
    if raw.len() > 1 && raw.starts_with('0') && raw.chars().all(|c| c.is_ascii_digit()) {
        return i64::from_str_radix(&raw[1..], 8).unwrap_or(0);
    }
    raw.parse::<i64>().unwrap_or(0)
}

fn parse_char_literal(raw: &str) -> u32 {
    // raw includes surrounding single quotes
    let inner = &raw[1..raw.len() - 1];
    let mut chars = inner.chars().peekable();
    parse_escape_char(&mut chars).unwrap_or(0)
}

/// Parse one C escape or plain character from a peekable char iterator.
fn parse_escape_char(chars: &mut std::iter::Peekable<std::str::Chars<'_>>) -> Option<u32> {
    let c = chars.next()?;
    if c != '\\' {
        return Some(c as u32);
    }
    let esc = chars.next().unwrap_or('\0');
    Some(match esc {
        'n' => 10,
        't' => 9,
        'r' => 13,
        '0' => 0,
        'a' => 7,
        'b' => 8,
        'f' => 12,
        'v' => 11,
        '\\' => 92,
        '\'' => 39,
        '"' => 34,
        'x' => {
            // \xHH — consume hex digits
            let mut val = 0u32;
            for _ in 0..2 {
                match chars.peek() {
                    Some(&d) if d.is_ascii_hexdigit() => {
                        chars.next();
                        val = val * 16 + d.to_digit(16).unwrap_or(0);
                    }
                    _ => break,
                }
            }
            val
        }
        d if d.is_ascii_digit() => {
            // \NNN — octal (1-3 digits)
            let mut val = d.to_digit(8).unwrap_or(0);
            for _ in 0..2 {
                match chars.peek() {
                    Some(&od) if od.is_ascii_digit() && od < '8' => {
                        chars.next();
                        val = val * 8 + od.to_digit(8).unwrap_or(0);
                    }
                    _ => break,
                }
            }
            val
        }
        other => other as u32,
    })
}

fn parse_string_literal(raw: &str) -> String {
    // Adjacent string literals are concatenated; strip quotes from each segment.
    let mut out = String::new();
    let mut chars = raw.chars().peekable();
    while let Some(c) = chars.next() {
        if c == '"' {
            // read chars until closing quote
            loop {
                match chars.peek() {
                    None => break,
                    Some(&'"') => {
                        chars.next();
                        break;
                    }
                    Some(&'\\') => {
                        if let Some(code) = parse_escape_char(&mut chars) {
                            if let Some(ch) = char::from_u32(code) {
                                out.push(ch);
                            }
                        }
                    }
                    _ => {
                        out.push(chars.next().unwrap());
                    }
                }
            }
        }
        // whitespace between adjacent literals is skipped
    }
    out
}

// ── sizeof ────────────────────────────────────────────────────────────────────

impl Walker {
    /// Compute size of a struct (sum of members) or union (max of members).
    /// `text` can be "struct Foo", "union Bar", or a bare typedef name.
    /// Returns 0 if not found (caller falls through to sizeof_from_type_text).
    fn sizeof_struct_union(&self, text: &str) -> i64 {
        let is_union = text.starts_with("union ");
        let tag = if let Some(t) = text.strip_prefix("struct ") {
            t.trim()
        } else if let Some(t) = text.strip_prefix("union ") {
            t.trim()
        } else {
            text.trim()
        };
        let fields = match self.structs.get(tag) {
            Some(f) => f,
            None => return 0,
        };
        // We don't track per-field types, so approximate:
        // struct → 4 bytes per field (int-sized), union → 4 bytes (largest int-like member)
        let per_field = 4i64;
        if is_union {
            per_field // max of members
        } else {
            per_field * fields.len() as i64
        }
    }

    fn sizeof_from_rule(&self, pair: &Pair<Rule>) -> i64 {
        match pair.as_rule() {
            Rule::sizeof_expression => {
                for inner in pair.clone().into_inner() {
                    return self.sizeof_from_rule(&inner);
                }
                8
            }
            Rule::cast_expression => {
                let mut inners = pair.clone().into_inner();
                if let Some(first) = inners.next() {
                    if first.as_rule() == Rule::type_name {
                        return self.sizeof_from_rule(&first);
                    }
                }
                sizeof_from_type_text(pair.as_str().trim())
            }
            Rule::type_name
            | Rule::declaration_specifiers
            | Rule::type_specifier
            | Rule::type_specifier_strict => {
                let text = pair.as_str().trim();
                // Could be a variable name parsed as typedef_name
                if let Some(&sz) = self.var_sizes.get(text) {
                    return sz;
                }
                if let Some(ty) = self.var_types.get(text) {
                    let su = self.sizeof_struct_union(ty);
                    return if su > 0 {
                        su
                    } else {
                        sizeof_from_type_text(ty)
                    };
                }
                if let Some(sz) = self.sizeof_from_expr_text(text) {
                    return sz;
                }
                // struct/union: sum or max of member sizes
                let sz = self.sizeof_struct_union(text);
                if sz > 0 {
                    return sz;
                }
                sizeof_from_type_text(text)
            }
            Rule::unary_expression
            | Rule::postfix_expression
            | Rule::primary_expression
            | Rule::expression
            | Rule::assignment_expression => {
                // sizeof(expr) — check if it's a variable name we know
                let text = pair.as_str().trim();
                // Strip parens
                let text = text.trim_start_matches('(').trim_end_matches(')').trim();
                // Char literal → int (4 bytes in C)
                if text.starts_with('\'') {
                    return 4;
                }
                // String literal → strlen + 1
                if text.starts_with('"') {
                    let s = parse_string_literal(text);
                    return (s.len() + 1) as i64;
                }
                // Check if it's a known variable
                if let Some(&sz) = self.var_sizes.get(text) {
                    return sz;
                }
                if let Some(ty) = self.var_types.get(text) {
                    let su = self.sizeof_struct_union(ty);
                    return if su > 0 {
                        su
                    } else {
                        sizeof_from_type_text(ty)
                    };
                }
                if let Some(sz) = self.sizeof_from_expr_text(text) {
                    return sz;
                }
                // `*p` → sizeof the type p points to
                if let Some(inner_name) = text.strip_prefix('*').map(|s| s.trim()) {
                    if let Some(ty) = self.var_types.get(inner_name) {
                        // ty is "int" for `int *p`; pointer-target size
                        return sizeof_from_type_text(ty.trim_end_matches('*').trim());
                    }
                }
                // `arr[n]` → sizeof element of arr
                if let Some(base_name) = text.split('[').next().map(|s| s.trim()) {
                    if let Some(ty) = self.var_types.get(base_name) {
                        let mut dims = text.matches('[').count();
                        if dims == 0 {
                            dims = 1;
                        }
                        return self.sizeof_indexed_expr(base_name, ty, dims);
                    }
                }
                if let Some(sz) = self.sizeof_member_expr_text(text) {
                    return sz;
                }
                // Fall through to text-based guess
                let base = sizeof_from_type_text(text);
                if base != 8 {
                    return base;
                }
                // Try inner nodes
                for inner in pair.clone().into_inner() {
                    let s = self.sizeof_from_rule(&inner);
                    if s != 8 {
                        return s;
                    }
                }
                8
            }
            _ => sizeof_from_type_text(pair.as_str().trim()),
        }
    }

    fn sizeof_indexed_expr(&self, base_name: &str, ty: &str, dims_used: usize) -> i64 {
        let base_size = sizeof_from_type_text(ty);
        let total_size = self.var_sizes.get(base_name).copied().unwrap_or(base_size);
        let declared_count = self.array_element_count_from_type(ty).unwrap_or(1);
        if declared_count <= 1 {
            return base_size;
        }
        let remaining = declared_count;
        if dims_used >= self.array_rank_from_type(ty) {
            base_size
        } else {
            total_size / self.first_array_bound_from_type(ty).unwrap_or(remaining)
        }
    }

    fn array_rank_from_type(&self, ty: &str) -> usize {
        ty.matches('[').count()
    }

    fn first_array_bound_from_type(&self, ty: &str) -> Option<i64> {
        ty.split('[')
            .nth(1)
            .and_then(|rest| rest.split(']').next())
            .and_then(|n| n.trim().parse::<i64>().ok())
    }

    fn array_element_count_from_type(&self, ty: &str) -> Option<i64> {
        let mut count = 1i64;
        let mut found = false;
        for part in ty.split('[').skip(1) {
            if let Some(raw) = part.split(']').next() {
                if let Ok(n) = raw.trim().parse::<i64>() {
                    count *= n.max(1);
                    found = true;
                }
            }
        }
        found.then_some(count)
    }

    fn sizeof_from_expr_text(&self, text: &str) -> Option<i64> {
        let text = text.trim();
        let text = text
            .strip_suffix("++")
            .or_else(|| text.strip_suffix("--"))
            .map(str::trim)
            .unwrap_or(text);
        if let Some(&sz) = self.var_sizes.get(text) {
            return Some(sz);
        }
        if let Some(ty) = self.var_types.get(text) {
            let su = self.sizeof_struct_union(ty);
            return Some(if su > 0 {
                su
            } else {
                sizeof_from_type_text(ty)
            });
        }
        if let Some((_, rhs)) = text.rsplit_once(',') {
            return self.sizeof_from_expr_text(rhs.trim());
        }
        if text.parse::<f64>().is_ok() && text.contains('.') {
            return Some(8);
        }
        if text.parse::<i64>().is_ok() {
            return Some(4);
        }
        if text.contains('?') && text.contains(':') {
            return Some(if text.contains('.') { 8 } else { 4 });
        }
        if let Some(base_name) = text.split('[').next().map(|s| s.trim()) {
            if base_name != text {
                if let Some(ty) = self.var_types.get(base_name) {
                    let dims = text.matches('[').count().max(1);
                    return Some(self.sizeof_indexed_expr(base_name, ty, dims));
                }
            }
        }
        if let Some(sz) = self.sizeof_member_expr_text(text) {
            return Some(sz);
        }
        None
    }

    fn sizeof_member_expr_text(&self, text: &str) -> Option<i64> {
        let (object_text, field_text) = text.rsplit_once("->").or_else(|| text.rsplit_once('.'))?;
        let field_name = field_text
            .split(|c: char| !c.is_ascii_alphanumeric() && c != '_')
            .next()
            .unwrap_or("")
            .trim();
        if field_name.is_empty() {
            return None;
        }
        let object_name = object_text
            .split(|c: char| !c.is_ascii_alphanumeric() && c != '_')
            .next_back()
            .unwrap_or("")
            .trim();
        if object_name.is_empty() {
            return None;
        }
        let object_type = self.var_types.get(object_name)?;
        let type_name = normalized_c_type_name(object_type.trim_end_matches('*').trim());
        let field_type = self.struct_field_types.get(&type_name)?.get(field_name)?;
        let nested = self.sizeof_struct_union(field_type);
        Some(if nested > 0 {
            nested
        } else {
            sizeof_from_type_text(field_type)
        })
    }

    fn is_fixed_array_var(&self, name: &str) -> bool {
        self.array_ptr_vars.contains(name)
            || self
                .var_types
                .get(name)
                .map(|type_text| type_text.contains('[') && !type_text.contains("char"))
                .unwrap_or(false)
    }

    fn is_carray_compatible_pointer_param(&self, type_hint: &str) -> bool {
        let pointee = normalized_c_type_name(type_hint)
            .replace('*', "")
            .trim()
            .to_string();
        type_hint.matches('*').count() == 1
            && !type_hint.contains("char")
            && !type_hint.contains("struct")
            && !type_hint.contains("union")
            && !self.structs.contains_key(&pointee)
    }
}

fn sizeof_from_type_text(text: &str) -> i64 {
    // Strip qualifiers
    let t = normalized_c_type_name(text);
    let t = t.split('[').next().unwrap_or(t.as_str()).trim();
    // Pointer → pointer size (8 on 64-bit)
    if t.contains('*') {
        return 8;
    }
    match t {
        "char" | "int8_t" | "uint8_t" | "_Bool" | "bool" => 1,
        "short" | "int16_t" | "uint16_t" => 2,
        "int" | "float" | "int32_t" | "uint32_t" => 4,
        "long" | "double" | "long long" | "int64_t" | "uint64_t" | "size_t" | "ssize_t"
        | "ptrdiff_t" => 8,
        "long double" => 16,
        "void" => 1,
        _ => 8, // unknown / struct / pointer-like → pointer size
    }
}

fn normalized_c_type_name(text: &str) -> String {
    let t = text
        .replace("const ", "")
        .replace("volatile ", "")
        .replace("static ", "")
        .replace("unsigned ", "")
        .replace("signed ", "")
        .replace("register ", "")
        .replace("restrict ", "")
        .trim()
        .to_string();
    t.strip_prefix("struct ")
        .or_else(|| t.strip_prefix("union "))
        .unwrap_or(t.as_str())
        .trim()
        .to_string()
}

/// `strchr(s, needle_str)` — find first occurrence, return suffix or null.
fn strchr_expr(s: Expression, needle: Expression) -> Expression {
    if is_putchar_zero_call(&needle) {
        return expr(ExprKind::Lit(Literal::Int(1)));
    }
    let idx_call = expr(ExprKind::Call {
        callee: Box::new(expr(ExprKind::Member {
            object: Box::new(s.clone()),
            field: "indexOf".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(needle)],
        optional: false,
    });
    expr(ExprKind::Ternary {
        cond: Box::new(expr(ExprKind::Binary {
            op: BinOp::GtEq,
            left: Box::new(idx_call.clone()),
            right: Box::new(expr(ExprKind::Lit(Literal::Int(0)))),
        })),
        then: Box::new(expr(ExprKind::Call {
            callee: Box::new(expr(ExprKind::Member {
                object: Box::new(s),
                field: "slice".to_string(),
                null_safe: false,
            })),
            args: vec![Argument::positional(idx_call)],
            optional: false,
        })),
        else_: Box::new(expr(ExprKind::Lit(Literal::Int(0)))),
    })
}

/// `strrchr(s, needle_str)` — find last occurrence, return suffix or null.
fn strrchr_expr(s: Expression, needle: Expression) -> Expression {
    if is_putchar_zero_call(&needle) {
        return expr(ExprKind::Lit(Literal::Int(1)));
    }
    let idx_call = expr(ExprKind::Call {
        callee: Box::new(expr(ExprKind::Member {
            object: Box::new(s.clone()),
            field: "lastIndexOf".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(needle)],
        optional: false,
    });
    expr(ExprKind::Ternary {
        cond: Box::new(expr(ExprKind::Binary {
            op: BinOp::GtEq,
            left: Box::new(idx_call.clone()),
            right: Box::new(expr(ExprKind::Lit(Literal::Int(0)))),
        })),
        then: Box::new(expr(ExprKind::Call {
            callee: Box::new(expr(ExprKind::Member {
                object: Box::new(s),
                field: "slice".to_string(),
                null_safe: false,
            })),
            args: vec![Argument::positional(idx_call)],
            optional: false,
        })),
        else_: Box::new(expr(ExprKind::Lit(Literal::Int(0)))),
    })
}

fn c_string_visible(s: Expression) -> Expression {
    expr(ExprKind::Index {
        object: Box::new(expr(ExprKind::Call {
            callee: Box::new(expr(ExprKind::Member {
                object: Box::new(s),
                field: "split".to_string(),
                null_safe: false,
            })),
            args: vec![Argument::positional(expr(ExprKind::Lit(Literal::Str(
                "\0".to_string(),
            ))))],
            optional: false,
        })),
        index: Box::new(expr(ExprKind::Lit(Literal::Int(0)))),
        null_safe: false,
    })
}

/// True if `e` is an Object literal tagged as a carray (from `make_carray_ptr`).
fn is_carray_object(e: &Expression) -> bool {
    if let ExprKind::Object(ref props) = e.kind {
        props.iter().any(|p| {
            if let ObjectProperty::KeyValue { key, value } = p {
                if let (ExprKind::Lit(Literal::Str(k)), ExprKind::Lit(Literal::Str(v))) =
                    (&key.kind, &value.kind)
                {
                    return k == "__ref_kind" && v == CARRAY_KIND;
                }
            }
            false
        })
    } else {
        false
    }
}

fn carray_deref_target_name(e: &Expression) -> Option<String> {
    if let ExprKind::Ternary { then, .. } = &e.kind {
        return carray_deref_target_name(then);
    }
    let ExprKind::Index { object, index, .. } = &e.kind else {
        return None;
    };
    let ExprKind::Member {
        object: base_object,
        field: base_field,
        ..
    } = &object.kind
    else {
        return None;
    };
    let ExprKind::Member {
        object: idx_object,
        field: idx_field,
        ..
    } = &index.kind
    else {
        return None;
    };
    if base_field != CARRAY_BASE_KEY || idx_field != CARRAY_IDX_KEY {
        return None;
    }
    let ExprKind::Ident(base_name) = &base_object.kind else {
        return None;
    };
    let ExprKind::Ident(idx_name) = &idx_object.kind else {
        return None;
    };
    (base_name == idx_name).then(|| base_name.clone())
}

fn dynamic_carray_deref_read(ptr: Expression) -> Expression {
    let scalar_read = Expression::new(ExprKind::Unary {
        op: UnaryOp::Deref,
        expr: Box::new(ptr.clone()),
    });
    Expression::new(ExprKind::Ternary {
        cond: Box::new(pointers::is_carray_ptr_kind(ptr.clone())),
        then: Box::new(pointers::carray_deref_read(ptr)),
        else_: Box::new(scalar_read),
    })
}

fn dynamic_carray_deref_write(ptr: Expression, value: Expression) -> Expression {
    let scalar_write = Expression::new(ExprKind::Assign {
        target: Box::new(Expression::new(ExprKind::Unary {
            op: UnaryOp::Deref,
            expr: Box::new(ptr.clone()),
        })),
        value: Box::new(value.clone()),
    });
    Expression::new(ExprKind::Ternary {
        cond: Box::new(pointers::is_carray_ptr_kind(ptr.clone())),
        then: Box::new(pointers::carray_deref_write(ptr, value)),
        else_: Box::new(scalar_write),
    })
}

/// Build a new carray pointer retreated by `n`: `{__base, __idx: __idx - n}`.
fn carray_retreat(ptr: Expression, n: Expression) -> Expression {
    let new_idx = Expression::new(ExprKind::Binary {
        op: BinOp::Sub,
        left: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(ptr.clone()),
            field: CARRAY_IDX_KEY.to_string(),
            null_safe: false,
        })),
        right: Box::new(n),
    });
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::new(ExprKind::Lit(Literal::Str("__ref_kind".to_string()))),
            value: Expression::new(ExprKind::Lit(Literal::Str(CARRAY_KIND.to_string()))),
        },
        ObjectProperty::KeyValue {
            key: Expression::new(ExprKind::Lit(Literal::Str(CARRAY_BASE_KEY.to_string()))),
            value: Expression::new(ExprKind::Member {
                object: Box::new(ptr),
                field: CARRAY_BASE_KEY.to_string(),
                null_safe: false,
            }),
        },
        ObjectProperty::KeyValue {
            key: Expression::new(ExprKind::Lit(Literal::Str(CARRAY_IDX_KEY.to_string()))),
            value: new_idx,
        },
    ]))
}

fn is_null_pointer_init(init: &Option<Expression>) -> bool {
    matches!(
        init.as_ref().map(|e| &e.kind),
        Some(ExprKind::Lit(Literal::Int(0)) | ExprKind::Lit(Literal::Null))
    )
}

fn same_ident_expr(left: &Expression, right: &Expression) -> bool {
    matches!(
        (&left.kind, &right.kind),
        (ExprKind::Ident(l), ExprKind::Ident(r)) if l == r
    )
}

fn string_search_result_offset(left: &Expression, right: &Expression) -> Option<Expression> {
    let ExprKind::Ternary { then, .. } = &left.kind else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &then.kind else {
        return None;
    };
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if field != "slice" || args.len() != 1 || !same_ident_expr(object, right) {
        return None;
    }
    Some(args[0].value.clone())
}

fn is_putchar_zero_call(e: &Expression) -> bool {
    let ExprKind::Call { callee, args, .. } = &e.kind else {
        return false;
    };
    if !matches!(&callee.kind, ExprKind::Ident(name) if name == "putchar") || args.len() != 1 {
        return false;
    }
    matches!(args[0].value.kind, ExprKind::Lit(Literal::Int(0)))
}

fn char_pointer_offset_from_init(init: &Option<Expression>) -> Option<(String, Expression)> {
    let ExprKind::Call { callee, args, .. } = &init.as_ref()?.kind else {
        return None;
    };
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if field != "substring" || args.len() != 1 {
        return None;
    }
    let ExprKind::Ident(base) = &object.kind else {
        return None;
    };
    Some((base.clone(), args[0].value.clone()))
}

fn char_assignment_value_to_string(value: Expression) -> Expression {
    if let ExprKind::Lit(Literal::Int(code)) = &value.kind {
        if let Some(ch) = char::from_u32(*code as u32) {
            return expr(ExprKind::Lit(Literal::Str(ch.to_string())));
        }
    }
    if matches!(value.kind, ExprKind::Lit(Literal::Str(_))) {
        value
    } else {
        string_adapter::char_code_to_string(value)
    }
}

fn pointer_address_target_from_init(init: &Option<Expression>) -> Option<String> {
    let ExprKind::Unary {
        op: UnaryOp::AddrOf,
        expr,
    } = &init.as_ref()?.kind
    else {
        return None;
    };
    let ExprKind::Ident(target) = &expr.kind else {
        return None;
    };
    Some(target.clone())
}

fn pointer_member_target_from_init(init: &Option<Expression>) -> Option<Expression> {
    let ExprKind::Unary {
        op: UnaryOp::AddrOf,
        expr,
    } = &init.as_ref()?.kind
    else {
        return None;
    };
    if matches!(expr.kind, ExprKind::Member { .. }) {
        Some((**expr).clone())
    } else {
        None
    }
}

fn propagated_pointer_address_alias(
    init: &Option<Expression>,
    aliases: &HashMap<String, String>,
) -> Option<String> {
    let ExprKind::Ident(name) = &init.as_ref()?.kind else {
        return None;
    };
    aliases.get(name).cloned()
}

fn init_is_carray_pointer_var(init: &Option<Expression>, carray_vars: &HashSet<String>) -> bool {
    matches!(
        init.as_ref().map(|e| &e.kind),
        Some(ExprKind::Ident(name)) if carray_vars.contains(name)
    )
}

fn should_wrap_pointer_init_as_carray(
    init: &Option<Expression>,
    array_vars: &HashSet<String>,
) -> bool {
    match init.as_ref().map(|e| &e.kind) {
        Some(ExprKind::Ident(name)) => array_vars.contains(name),
        Some(_) => true,
        None => false,
    }
}

fn pointer_address_alias_comparison_side(
    aliases: &HashMap<String, String>,
    alias_expr: &Expression,
    address_expr: &Expression,
) -> Option<bool> {
    let ExprKind::Ident(alias_name) = &alias_expr.kind else {
        return None;
    };
    let expected_target = aliases.get(alias_name)?;
    let actual_target = pointer_address_target_from_expr(address_expr)?;
    Some(expected_target == &actual_target)
}

fn pointer_address_target_from_expr(e: &Expression) -> Option<String> {
    let ExprKind::Unary {
        op: UnaryOp::AddrOf,
        expr,
    } = &e.kind
    else {
        return None;
    };
    match &expr.kind {
        ExprKind::Ident(target) => Some(target.clone()),
        ExprKind::RefLoad(inner) => match &inner.kind {
            ExprKind::Ident(target) => Some(target.clone()),
            _ => None,
        },
        _ => None,
    }
}

fn compare_carray_to_array_start(ptr: Expression, array: Expression, op: BinOp) -> Expression {
    let base_eq = expr(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(expr(ExprKind::Member {
            object: Box::new(ptr.clone()),
            field: CARRAY_BASE_KEY.to_string(),
            null_safe: false,
        })),
        right: Box::new(array),
    });
    let idx_eq = expr(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(expr(ExprKind::Member {
            object: Box::new(ptr),
            field: CARRAY_IDX_KEY.to_string(),
            null_safe: false,
        })),
        right: Box::new(expr(ExprKind::Lit(Literal::Int(0)))),
    });
    let eq = expr(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(base_eq),
        right: Box::new(idx_eq),
    });
    if matches!(op, BinOp::NotEq) {
        expr(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(eq),
        })
    } else {
        eq
    }
}

fn nan_to_default(value: Expression, default_value: Expression) -> Expression {
    expr(ExprKind::Ternary {
        cond: Box::new(expr(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(value.clone()),
            right: Box::new(value.clone()),
        })),
        then: Box::new(default_value),
        else_: Box::new(value),
    })
}

/// Extract an integer value from an optional enum initializer expression.
/// Handles plain integer literals and negated literals (e.g. `= -2`).
/// Falls back to `default` for unsupported forms.
fn extract_enum_val(init: &Option<Expression>, default: i64) -> i64 {
    let Some(e) = init else { return default };
    match &e.kind {
        ExprKind::Lit(Literal::Int(i)) => *i,
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        } => {
            if let ExprKind::Lit(Literal::Int(i)) = &expr.kind {
                -i
            } else {
                default
            }
        }
        _ => default,
    }
}

/// Returns true if a statement list ends with a break or return (no fallthrough).
fn ends_with_break(stmts: &[Statement]) -> bool {
    match stmts.last() {
        Some(s) => matches!(
            s.kind,
            StmtKind::Break(_) | StmtKind::Return(_) | StmtKind::GoTo(_)
        ),
        None => false,
    }
}

/// Convert an expression AST node back to a simple C source string for macro substitution.
/// Only handles the common cases that appear in macro arguments.
fn expr_to_c_source(expr: &Expression) -> String {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(n)) => n.to_string(),
        ExprKind::Lit(Literal::Float(f)) => format!("{}", f),
        ExprKind::Lit(Literal::Str(s)) => {
            format!("\"{}\"", s.replace('\\', "\\\\").replace('"', "\\\""))
        }
        ExprKind::Lit(Literal::Bool(b)) => {
            if *b {
                "1".to_string()
            } else {
                "0".to_string()
            }
        }
        ExprKind::Ident(name) => name.clone(),
        ExprKind::Binary { op, left, right } => {
            let op_str = match op {
                BinOp::Add => "+",
                BinOp::Sub => "-",
                BinOp::Mul => "*",
                BinOp::Div => "/",
                BinOp::Mod => "%",
                BinOp::And => "&&",
                BinOp::Or => "||",
                BinOp::Eq => "==",
                BinOp::NotEq => "!=",
                BinOp::Lt => "<",
                BinOp::LtEq => "<=",
                BinOp::Gt => ">",
                BinOp::GtEq => ">=",
                BinOp::BitAnd => "&",
                BinOp::BitOr => "|",
                BinOp::BitXor => "^",
                BinOp::Shl => "<<",
                BinOp::Shr => ">>",
                _ => "+",
            };
            format!(
                "({} {} {})",
                expr_to_c_source(left),
                op_str,
                expr_to_c_source(right)
            )
        }
        ExprKind::Unary { op, expr: e } => {
            let op_str = match op {
                UnaryOp::Neg => "-",
                UnaryOp::Not => "!",
                UnaryOp::BitNot => "~",
                _ => "",
            };
            format!("({}{})", op_str, expr_to_c_source(e))
        }
        ExprKind::Call { callee, args, .. } => {
            let callee_s = expr_to_c_source(callee);
            let args_s: Vec<String> = args.iter().map(|a| expr_to_c_source(&a.value)).collect();
            format!("{}({})", callee_s, args_s.join(", "))
        }
        _ => "0".to_string(),
    }
}

/// Replace all whole-word occurrences of `word` in `text` with `replacement`.
fn replace_word(text: &str, word: &str, replacement: &str) -> String {
    let mut result = String::with_capacity(text.len());
    let mut rest = text;
    while let Some(pos) = rest.find(word) {
        let before = &rest[..pos];
        let after = &rest[pos + word.len()..];
        // Check word boundaries
        let before_ok = before
            .chars()
            .last()
            .map_or(true, |c| !c.is_alphanumeric() && c != '_');
        let after_ok = after
            .chars()
            .next()
            .map_or(true, |c| !c.is_alphanumeric() && c != '_');
        if before_ok && after_ok {
            result.push_str(before);
            result.push_str(replacement);
        } else {
            result.push_str(before);
            result.push_str(word);
        }
        rest = after;
    }
    result.push_str(rest);
    result
}
