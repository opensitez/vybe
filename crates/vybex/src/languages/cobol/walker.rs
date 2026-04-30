//! COBOL walker — pest `Pair<Rule>` → `vybex::ast::Module`.
//!
//! Walks the parse tree produced by `grammar.pest` into the common AST.
//! Once this returns a `Module`, the rest of the compilation pipeline
//! (compile_class / compile_expression / etc.) is shared with every
//! other vybex language and works without any COBOL-specific knowledge.
//!
//! ## COBOL semantics normalised by the walker
//!
//! - **PERFORM UNTIL** means "loop while condition is FALSE". The walker
//!   negates the condition so the downstream compiler sees a regular While.
//!
//! - **1-indexed arrays**: COBOL arrays start at 1. The walker subtracts 1
//!   from all subscript indices.
//!
//! - **Paragraphs → FunctionDecl**: Each paragraph in the procedure
//!   division becomes a zero-parameter FunctionDecl.
//!
//! - **Data items → VarDecl**: Level 01/77 at the top become VarDecl.
//!   Group items create struct-like Object initialisers. OCCURS creates
//!   Array initialisers.
//!
//! - **Figurative constants**: SPACES → " ", ZEROS → 0, etc.
//!
//! - **88-level conditions**: Become Const members on the parent or
//!   standalone VarDecl constants.
//!
//! - **Period terminators**: Skipped — they are grammar structure, not AST
//!   nodes.
//!
//! - **EXEC SQL/CICS/DLI blocks**: The raw text between EXEC and END-EXEC
//!   is passed as a string literal argument to `__exec_sql` / `__exec_cics`
//!   / `__exec_dli` helper calls.

use pest::Parser;
use pest::iterators::{Pair, Pairs};
use crate::ast::*;
use super::{CobolParser, Rule};

// ════════════════════════════════════════════════════════════════════════════
// Keyword filter
// ════════════════════════════════════════════════════════════════════════════

/// Returns true for `kw_*` token rules. Pest preserves atomic rule nodes as
/// siblings inside their parent rule's parse tree, so without this filter
/// the keyword tokens leak into walker positional indexing.
fn is_kw(r: Rule) -> bool {
    use Rule::*;
    matches!(r,
        kw_identification | kw_environment | kw_configuration | kw_data
        | kw_procedure | kw_division | kw_section | kw_program_id | kw_class_id
        | kw_interface_id | kw_method_id | kw_author | kw_date_written
        | kw_special_names | kw_repository | kw_input_output | kw_file_control
        | kw_decimal_point | kw_currency | kw_comma | kw_working_storage
        | kw_local_storage | kw_linkage | kw_file | kw_screen | kw_pic
        | kw_picture | kw_value | kw_occurs | kw_times | kw_depending
        | kw_redefines | kw_usage | kw_indexed | kw_filler | kw_blank
        | kw_justified | kw_just | kw_synchronized | kw_sync | kw_global
        | kw_external | kw_binary | kw_comp | kw_comp_3 | kw_comp_5
        | kw_display_usage | kw_packed_decimal | kw_pointer | kw_index
        | kw_float_long | kw_float_short | kw_national | kw_boolean
        | kw_fd | kw_sd | kw_record | kw_block | kw_contains | kw_characters
        | kw_character | kw_label | kw_standard | kw_omitted | kw_spaces
        | kw_space | kw_zeros | kw_zeroes | kw_zero | kw_low_values
        | kw_low_value | kw_high_values | kw_high_value | kw_quotes
        | kw_quote | kw_nulls | kw_null | kw_all | kw_display | kw_accept
        | kw_move | kw_add | kw_subtract | kw_multiply | kw_divide
        | kw_compute | kw_if | kw_else | kw_then | kw_evaluate | kw_when
        | kw_other | kw_perform | kw_until | kw_varying | kw_thru
        | kw_through | kw_string | kw_unstring | kw_inspect | kw_tallying
        | kw_replacing | kw_converting | kw_leading | kw_trailing | kw_first
        | kw_initial | kw_call | kw_using | kw_returning | kw_initialize
        | kw_set | kw_go | kw_to | kw_stop | kw_run | kw_goback
        | kw_continue | kw_raise | kw_exception | kw_json | kw_generate
        | kw_parse | kw_open | kw_close | kw_read | kw_write | kw_rewrite
        | kw_delete | kw_start | kw_sort | kw_merge | kw_search | kw_copy
        | kw_invoke | kw_validate | kw_free | kw_allocate | kw_typedef
        | kw_exit | kw_not | kw_and | kw_or | kw_true | kw_false | kw_any
        | kw_with | kw_test | kw_before | kw_after | kw_async | kw_giving
        | kw_from | kw_by | kw_into | kw_on | kw_size | kw_error
        | kw_rounded | kw_remainder | kw_corresponding | kw_corr
        | kw_delimited | kw_delimiter | kw_count | kw_overflow | kw_input
        | kw_output | kw_extend | kw_i_o | kw_ascending | kw_descending
        | kw_key | kw_at | kw_end | kw_invalid | kw_next | kw_page
        | kw_advancing | kw_lines | kw_line | kw_upon | kw_no | kw_numeric
        | kw_alphabetic | kw_alphabetic_lower | kw_alphabetic_upper
        | kw_positive | kw_negative | kw_equal | kw_greater | kw_less
        | kw_than | kw_not_less | kw_not_greater | kw_inherits
        | kw_implements | kw_factory | kw_object | kw_new | kw_self
        | kw_override | kw_property | kw_get | kw_is | kw_as | kw_wait
        | kw_for | kw_unit | kw_lock | kw_unlock | kw_yield | kw_suspend
        | kw_exec | kw_sql | kw_cics | kw_dli | kw_end_exec | kw_date
        | kw_time | kw_day | kw_day_of_week | kw_command_line | kw_console
        | kw_select | kw_assign | kw_organization | kw_sequential
        | kw_relative | kw_file_status | kw_access | kw_mode | kw_random
        | kw_dynamic | kw_alternate | kw_duplicates | kw_up | kw_down
        | kw_reference | kw_content | kw_alphanumeric | kw_sign | kw_separate
        | kw_end_if | kw_end_evaluate | kw_end_perform | kw_end_call
        | kw_end_read | kw_end_write | kw_end_rewrite | kw_end_delete
        | kw_end_start | kw_end_string | kw_end_unstring | kw_end_search
        | kw_end_add | kw_end_subtract | kw_end_multiply | kw_end_divide
        | kw_end_compute | kw_end_class | kw_end_method | kw_end_factory
        | kw_end_object | kw_end_interface | kw_end_validate | kw_end_json
        | kw_end_program | kw_name | kw_class | kw_paragraph | kw_program
        | kw_method | kw_cycle | kw_when_ | kw_left | kw_right | kw_function
        | kw_length | kw_upper_case | kw_lower_case | kw_trim | kw_reverse
        | kw_current_date | kw_max | kw_min | kw_mod | kw_rem | kw_numval
        | kw_numval_c | kw_substitute | kw_sqrt | kw_sum | kw_integer
        | kw_abs | kw_ord | kw_char | kw_floor | kw_ceiling | kw_power
        | kw_log | kw_log10 | kw_exp | kw_sin | kw_cos | kw_tan | kw_asin
        | kw_acos | kw_atan | kw_mean | kw_median | kw_variance
        | kw_concatenate | kw_when_compiled | kw_test_numval
        | kw_date_of_integer | kw_integer_of_date | kw_day_of_integer
        | kw_annuity | kw_present_value | kw_formatted_date
        | kw_formatted_time
        | kw_also | kw_in | kw_of
    )
}

/// `pair.into_inner()` with `kw_*` siblings stripped.
fn inner_nokw(pair: Pair<Rule>) -> Vec<Pair<Rule>> {
    pair.into_inner()
        .filter(|p| !is_kw(p.as_rule()) && !matches!(p.as_rule(), Rule::period))
        .collect()
}

/// Filter an existing iterator to skip kw and period.
fn filter_nokw(pairs: Pairs<Rule>) -> Vec<Pair<Rule>> {
    pairs
        .filter(|p| !is_kw(p.as_rule()) && !matches!(p.as_rule(), Rule::period))
        .collect()
}

// ════════════════════════════════════════════════════════════════════════════
// Span
// ════════════════════════════════════════════════════════════════════════════

fn to_span(pair: &Pair<Rule>) -> Span {
    let s = pair.as_span();
    let (start_line, start_col) = s.start_pos().line_col();
    let (end_line, end_col) = s.end_pos().line_col();
    Span {
        start_line: start_line as u32,
        start_col: start_col as u32,
        end_line: end_line as u32,
        end_col: end_col as u32,
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Top-level entry point
// ════════════════════════════════════════════════════════════════════════════

pub fn parse(source: &str) -> Result<Module, String> {
    let mut pairs = CobolParser::parse(Rule::program, source)
        .map_err(|e| format!("COBOL parse error: {}", e))?;
    let program = pairs.next().ok_or("empty parse")?;

    let mut module = Module {
        name: String::new(),
        language: Lang::Cobol,
        body: Vec::new(),
        imports: Vec::new(),
    };

    for pair in program.into_inner() {
        match pair.as_rule() {
            Rule::identification_division => {
                walk_identification_division(pair, &mut module)?;
            }
            Rule::environment_division => {
                // Environment division is mostly declarative config.
                // We extract file SELECT entries as comments / empty stmts.
                walk_environment_division(pair, &mut module)?;
            }
            Rule::data_division => {
                walk_data_division(pair, &mut module.body)?;
            }
            Rule::procedure_division => {
                walk_procedure_division(pair, &mut module.body)?;
            }
            Rule::nested_program => {
                walk_nested_program(pair, &mut module.body)?;
            }
            Rule::EOI => {}
            _ => {}
        }
    }

    Ok(module)
}

// ════════════════════════════════════════════════════════════════════════════
// Identification Division
// ════════════════════════════════════════════════════════════════════════════

fn walk_identification_division(pair: Pair<Rule>, module: &mut Module) -> Result<(), String> {
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::program_id_paragraph => {
                let parts = inner_nokw(child);
                if let Some(name_pair) = parts.into_iter().find(|p| p.as_rule() == Rule::ident_name) {
                    module.name = name_pair.as_str().to_string();
                }
            }
            Rule::class_id_paragraph => {
                walk_class_id(child, &mut module.body)?;
            }
            Rule::interface_id_paragraph => {
                walk_interface_id(child, &mut module.body)?;
            }
            Rule::id_optional_paragraph => {
                // AUTHOR, DATE-WRITTEN — skip
            }
            _ => {}
        }
    }
    Ok(())
}

// ════════════════════════════════════════════════════════════════════════════
// Environment Division
// ════════════════════════════════════════════════════════════════════════════

fn walk_environment_division(pair: Pair<Rule>, _module: &mut Module) -> Result<(), String> {
    // Environment division contains configuration and file control.
    // For now we skip these — they affect runtime semantics but the
    // common AST does not have a direct representation for file
    // organisation modes, special-names, etc.  File I/O statements
    // in the procedure division emit host calls directly.
    for _child in pair.into_inner() {
        // configuration_section, input_output_section — skipped
    }
    Ok(())
}

// ════════════════════════════════════════════════════════════════════════════
// Data Division
// ════════════════════════════════════════════════════════════════════════════

fn walk_data_division(pair: Pair<Rule>, body: &mut Vec<Statement>) -> Result<(), String> {
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::file_section => {
                walk_file_section(child, body)?;
            }
            Rule::working_storage_section
            | Rule::local_storage_section
            | Rule::linkage_section
            | Rule::screen_section => {
                walk_storage_section(child, body)?;
            }
            _ => {}
        }
    }
    Ok(())
}

fn walk_file_section(pair: Pair<Rule>, body: &mut Vec<Statement>) -> Result<(), String> {
    for child in pair.into_inner() {
        if child.as_rule() == Rule::file_description {
            walk_file_description(child, body)?;
        }
    }
    Ok(())
}

fn walk_file_description(pair: Pair<Rule>, body: &mut Vec<Statement>) -> Result<(), String> {
    // FD/SD name fd_clause* . data_item*
    // We skip FD clauses and just emit data items.
    for child in pair.into_inner() {
        if child.as_rule() == Rule::data_item {
            walk_data_item(child, body)?;
        }
    }
    Ok(())
}

fn walk_storage_section(pair: Pair<Rule>, body: &mut Vec<Statement>) -> Result<(), String> {
    for child in pair.into_inner() {
        if child.as_rule() == Rule::data_item {
            walk_data_item(child, body)?;
        }
    }
    Ok(())
}

// ── Data Items ─────────────────────────────────────────────────────────────

fn walk_data_item(pair: Pair<Rule>, body: &mut Vec<Statement>) -> Result<(), String> {
    let span = to_span(&pair);
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::level_88_item => {
                // 88-level condition → constant declaration
                let stmt = walk_level_88(child)?;
                body.push(Statement::with_span(stmt, span));
            }
            Rule::regular_data_item => {
                walk_regular_data_item(child, body)?;
            }
            Rule::data_item => {
                // Nested data items (children of a group)
                walk_data_item(child, body)?;
            }
            _ => {}
        }
    }
    Ok(())
}

fn walk_level_88(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let parts = inner_nokw(pair);
    let mut name = String::new();
    let mut value_expr = Expression::bool(true);

    for p in parts {
        match p.as_rule() {
            Rule::ident_name => {
                name = p.as_str().to_string();
            }
            Rule::literal => {
                value_expr = walk_literal(p)?;
            }
            _ => {}
        }
    }

    Ok(StmtKind::VarDecl {
        kind: VarDeclKind::Const,
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name),
            type_hint: None,
            init: Some(value_expr),
            array_bounds: None,
            with_events: false,
        }],
    })
}

fn walk_regular_data_item(pair: Pair<Rule>, body: &mut Vec<Statement>) -> Result<(), String> {
    let span = to_span(&pair);
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut _level: u32 = 0;
    let mut name = String::new();
    let mut pic_str: Option<String> = None;
    let mut init_value: Option<Expression> = None;
    let mut occurs_count: Option<Expression> = None;
    let mut is_filler = false;
    let mut nested_items: Vec<Pair<Rule>> = Vec::new();

    for child in children {
        match child.as_rule() {
            Rule::level_number => {
                _level = child.as_str().trim().parse::<u32>().unwrap_or(0);
            }
            Rule::ident_name | Rule::ident_or_keyword => {
                // Grammar rule `regular_data_item` emits the data
                // item's name via `ident_or_keyword` (allows names
                // that match COBOL keywords like `STATUS`). We
                // match both here so top-level AND nested items
                // populate `name` correctly. Historical bug: only
                // `ident_name` was matched, which silently empty'd
                // the name and caused the OCCURS branch in
                // `collect_group_children` to be short-circuited
                // via the FILLER path.
                if name.is_empty() {
                    name = child.as_str().to_string();
                }
            }
            Rule::kw_filler => {
                is_filler = true;
                name = "FILLER".to_string();
            }
            Rule::data_clause => {
                walk_data_clause(child, &mut pic_str, &mut init_value, &mut occurs_count)?;
            }
            Rule::level_88_item => {
                // 88-level children — emit as separate constant declarations
                let stmt = walk_level_88(child)?;
                body.push(Statement::with_span(stmt, span));
            }
            Rule::data_item => {
                nested_items.push(child);
            }
            Rule::period => {}
            _ => {
                if is_kw(child.as_rule()) { continue; }
            }
        }
    }

    // Skip FILLER items (they are padding)
    if is_filler {
        // Still process nested items
        for nested in nested_items {
            walk_data_item(nested, body)?;
        }
        return Ok(());
    }

    // Determine type hint from PIC
    let type_hint = pic_str.as_ref().map(|p| pic_to_type_hint(p));

    // Determine initial value
    let init = if !nested_items.is_empty() {
        // Group item → Object initialiser with child fields
        let mut props = Vec::new();
        let mut child_stmts = Vec::new();
        for nested in nested_items {
            // Walk nested item and collect as object property
            collect_group_children(nested, &mut props, &mut child_stmts)?;
        }
        // Also emit any 88-level constants that were in child_stmts
        for s in child_stmts {
            body.push(s);
        }
        if props.is_empty() {
            init_value
        } else {
            Some(Expression::new(ExprKind::Object(props)))
        }
    } else if let Some(count_expr) = occurs_count {
        // OCCURS → array initialiser
        let element_init = init_value.clone().unwrap_or_else(|| {
            default_value_for_pic(&pic_str)
        });
        Some(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("Array")),
            args: vec![
                Argument::positional(count_expr),
                Argument::positional(element_init),
            ],
            optional: false,
        }))
    } else {
        init_value.or_else(|| Some(default_value_for_pic(&pic_str)))
    };

    if name.is_empty() {
        return Ok(());
    }

    let stmt = StmtKind::VarDecl {
        kind: VarDeclKind::Dim,
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name),
            type_hint,
            init,
            array_bounds: None,
            with_events: false,
        }],
    };
    body.push(Statement::with_span(stmt, span));
    Ok(())
}

fn walk_data_clause(
    pair: Pair<Rule>,
    pic_str: &mut Option<String>,
    init_value: &mut Option<Expression>,
    occurs_count: &mut Option<Expression>,
) -> Result<(), String> {
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::pic_clause => {
                for p in child.into_inner() {
                    if p.as_rule() == Rule::pic_string {
                        *pic_str = Some(p.as_str().to_string());
                    }
                }
            }
            Rule::value_clause => {
                let parts = inner_nokw(child);
                for p in parts {
                    if p.as_rule() == Rule::literal {
                        *init_value = Some(walk_literal(p)?);
                    }
                }
            }
            Rule::occurs_clause => {
                let parts = inner_nokw(child);
                for p in parts {
                    if p.as_rule() == Rule::number_literal {
                        *occurs_count = Some(parse_number_literal(p.as_str()));
                        break;
                    }
                    if p.as_rule() == Rule::level_number {
                        *occurs_count = Some(parse_number_literal(p.as_str()));
                        break;
                    }
                }
            }
            Rule::redefines_clause | Rule::usage_clause | Rule::blank_clause
            | Rule::justified_clause | Rule::sign_clause | Rule::sync_clause
            | Rule::global_clause | Rule::external_clause | Rule::national_clause => {
                // These affect storage layout but not the AST structure
            }
            _ => {}
        }
    }
    Ok(())
}

/// Collect children of a group data item as ObjectProperty entries.
fn collect_group_children(
    pair: Pair<Rule>,
    props: &mut Vec<ObjectProperty>,
    extra_stmts: &mut Vec<Statement>,
) -> Result<(), String> {
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::level_88_item => {
                let stmt = walk_level_88(child)?;
                extra_stmts.push(Statement::new(stmt));
            }
            Rule::regular_data_item => {
                let children: Vec<Pair<Rule>> = child.into_inner().collect();
                let mut field_name = String::new();
                let mut field_pic: Option<String> = None;
                let mut field_init: Option<Expression> = None;
                let mut field_occurs: Option<Expression> = None;
                let mut sub_items: Vec<Pair<Rule>> = Vec::new();

                for c in children {
                    match c.as_rule() {
                        Rule::level_number => {}
                        Rule::ident_name | Rule::ident_or_keyword => {
                            // See the equivalent fix in
                            // `walk_regular_data_item` — the grammar
                            // emits `ident_or_keyword` for data-item
                            // names; matching only `ident_name` was
                            // a silent no-op that broke nested
                            // OCCURS handling.
                            if field_name.is_empty() {
                                field_name = c.as_str().to_string();
                            }
                        }
                        Rule::kw_filler => {
                            field_name = "FILLER".to_string();
                        }
                        Rule::data_clause => {
                            walk_data_clause(c, &mut field_pic, &mut field_init, &mut field_occurs)?;
                        }
                        Rule::level_88_item => {
                            let stmt = walk_level_88(c)?;
                            extra_stmts.push(Statement::new(stmt));
                        }
                        Rule::data_item => {
                            sub_items.push(c);
                        }
                        Rule::period => {}
                        _ => {}
                    }
                }

                if field_name == "FILLER" || field_name.is_empty() {
                    // Recurse into sub-items even for FILLER
                    for si in sub_items {
                        collect_group_children(si, props, extra_stmts)?;
                    }
                    continue;
                }

                let value = if !sub_items.is_empty() {
                    let mut sub_props = Vec::new();
                    for si in sub_items {
                        collect_group_children(si, &mut sub_props, extra_stmts)?;
                    }
                    Expression::new(ExprKind::Object(sub_props))
                } else if let Some(count_expr) = field_occurs {
                    let element_init = field_init.unwrap_or_else(|| default_value_for_pic(&field_pic));
                    eprintln!("[D1 WALKER] nested OCCURS path emitting Array(count, init) for field");
                    Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("Array")),
                        args: vec![
                            Argument::positional(count_expr),
                            Argument::positional(element_init),
                        ],
                        optional: false,
                    })
                } else {
                    field_init.unwrap_or_else(|| default_value_for_pic(&field_pic))
                };

                props.push(ObjectProperty::KeyValue {
                    key: Expression::string(&field_name),
                    value,
                });
            }
            Rule::data_item => {
                collect_group_children(child, props, extra_stmts)?;
            }
            _ => {}
        }
    }
    Ok(())
}

/// Determine a type hint from a PIC string.
fn pic_to_type_hint(pic: &str) -> String {
    let upper = pic.to_uppercase();
    if upper.starts_with('X') || upper.starts_with('A') {
        "String".to_string()
    } else if upper.contains('V') || upper.contains('.') {
        "Double".to_string()
    } else {
        "Integer".to_string()
    }
}

/// Default value for a PIC type (spaces for alpha, 0 for numeric).
fn default_value_for_pic(pic: &Option<String>) -> Expression {
    if let Some(p) = pic {
        let upper = p.to_uppercase();
        if upper.starts_with('X') || upper.starts_with('A') {
            Expression::string(" ")
        } else {
            Expression::int(0)
        }
    } else {
        Expression::null()
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Procedure Division
// ════════════════════════════════════════════════════════════════════════════

fn walk_procedure_division(pair: Pair<Rule>, body: &mut Vec<Statement>) -> Result<(), String> {
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::using_clause | Rule::returning_clause => {
                // Procedure division USING / RETURNING — parameters for the
                // main program.  These are rarely used in practice.
            }
            Rule::statement_list => {
                walk_statement_list(child, body)?;
            }
            Rule::paragraph => {
                walk_paragraph(child, body)?;
            }
            Rule::period => {}
            _ => {
                if !is_kw(child.as_rule()) {
                    // Skip unexpected tokens
                }
            }
        }
    }
    Ok(())
}

fn walk_paragraph(pair: Pair<Rule>, body: &mut Vec<Statement>) -> Result<(), String> {
    let span = to_span(&pair);
    let parts = filter_nokw(pair.into_inner());

    let mut name = String::new();
    let mut para_body = Vec::new();

    for p in parts {
        match p.as_rule() {
            Rule::paragraph_name => {
                name = p.as_str().to_string();
            }
            Rule::statement_list => {
                walk_statement_list(p, &mut para_body)?;
            }
            _ => {}
        }
    }

    if name.is_empty() {
        return Ok(());
    }

    body.push(Statement::with_span(
        StmtKind::FunctionDecl {
            name,
            params: Vec::new(),
            return_type: None,
            body: para_body,
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async: false,
            is_generator: false,
            is_sub: true,
        },
        span,
    ));
    Ok(())
}

fn walk_statement_list(pair: Pair<Rule>, body: &mut Vec<Statement>) -> Result<(), String> {
    for child in pair.into_inner() {
        let rule = child.as_rule();
        if matches!(rule, Rule::period) || is_kw(rule) {
            continue;
        }
        if let Some(stmt) = walk_statement(child)? {
            body.push(stmt);
        }
    }
    Ok(())
}

// ════════════════════════════════════════════════════════════════════════════
// Statements
// ════════════════════════════════════════════════════════════════════════════

fn walk_statement(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    let span = to_span(&pair);
    let rule = pair.as_rule();

    let kind = match rule {
        // ── DISPLAY ─────────────────────────────────────────────────────
        Rule::display_stmt => {
            let mut exprs = Vec::new();
            for child in pair.into_inner() {
                match child.as_rule() {
                    Rule::expression => {
                        exprs.push(walk_expression(child)?);
                    }
                    _ => {}
                }
            }
            StmtKind::Echo(exprs)
        }

        // ── ACCEPT ──────────────────────────────────────────────────────
        Rule::accept_stmt => {
            walk_accept_stmt(pair)?
        }

        // ── MOVE ────────────────────────────────────────────────────────
        Rule::move_stmt => {
            walk_move_stmt(pair)?
        }

        // ── ADD ─────────────────────────────────────────────────────────
        Rule::add_stmt => {
            walk_add_stmt(pair)?
        }

        // ── SUBTRACT ────────────────────────────────────────────────────
        Rule::subtract_stmt => {
            walk_subtract_stmt(pair)?
        }

        // ── MULTIPLY ────────────────────────────────────────────────────
        Rule::multiply_stmt => {
            walk_multiply_stmt(pair)?
        }

        // ── DIVIDE ──────────────────────────────────────────────────────
        Rule::divide_stmt => {
            walk_divide_stmt(pair)?
        }

        // ── COMPUTE ─────────────────────────────────────────────────────
        Rule::compute_stmt => {
            walk_compute_stmt(pair)?
        }

        // ── IF ──────────────────────────────────────────────────────────
        Rule::if_stmt => {
            walk_if_stmt(pair)?
        }

        // ── EVALUATE ────────────────────────────────────────────────────
        Rule::evaluate_stmt => {
            walk_evaluate_stmt(pair)?
        }

        // ── PERFORM ─────────────────────────────────────────────────────
        Rule::perform_stmt => {
            walk_perform_stmt(pair)?
        }

        // ── STRING ──────────────────────────────────────────────────────
        Rule::string_stmt => {
            walk_string_stmt(pair)?
        }

        // ── UNSTRING ────────────────────────────────────────────────────
        Rule::unstring_stmt => {
            walk_unstring_stmt(pair)?
        }

        // ── INSPECT ─────────────────────────────────────────────────────
        Rule::inspect_stmt => {
            walk_inspect_stmt(pair)?
        }

        // ── CALL ────────────────────────────────────────────────────────
        Rule::call_stmt => {
            walk_call_stmt(pair)?
        }

        // ── INITIALIZE ──────────────────────────────────────────────────
        Rule::initialize_stmt => {
            walk_initialize_stmt(pair)?
        }

        // ── SET ─────────────────────────────────────────────────────────
        Rule::set_stmt => {
            walk_set_stmt(pair)?
        }

        // ── GO TO ───────────────────────────────────────────────────────
        Rule::go_to_stmt => {
            let parts = inner_nokw(pair);
            let name = parts.into_iter()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            // GO TO paragraph → call the paragraph
            StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(&name)),
                args: Vec::new(),
                optional: false,
            }))
        }

        // ── STOP RUN ────────────────────────────────────────────────────
        Rule::stop_run_stmt => {
            StmtKind::Return(None)
        }

        // ── GOBACK ──────────────────────────────────────────────────────
        Rule::goback_stmt => {
            StmtKind::Return(None)
        }

        // ── CONTINUE ────────────────────────────────────────────────────
        Rule::continue_stmt => {
            StmtKind::Empty
        }

        // ── RAISE ───────────────────────────────────────────────────────
        Rule::raise_stmt => {
            let parts = inner_nokw(pair);
            let expr = parts.into_iter()
                .find(|p| p.as_rule() == Rule::expression)
                .map(|p| walk_expression(p))
                .transpose()?;
            StmtKind::Throw { expr, cause: None }
        }

        // ── JSON GENERATE / PARSE ───────────────────────────────────────
        Rule::json_stmt => {
            walk_json_stmt(pair)?
        }

        // ── OPEN ────────────────────────────────────────────────────────
        Rule::open_stmt => {
            walk_open_stmt(pair)?
        }

        // ── CLOSE ───────────────────────────────────────────────────────
        Rule::close_stmt => {
            walk_close_stmt(pair)?
        }

        // ── READ ────────────────────────────────────────────────────────
        Rule::read_stmt => {
            walk_read_stmt(pair)?
        }

        // ── WRITE ───────────────────────────────────────────────────────
        Rule::write_stmt => {
            walk_write_stmt(pair)?
        }

        // ── REWRITE ─────────────────────────────────────────────────────
        Rule::rewrite_stmt => {
            walk_rewrite_stmt(pair)?
        }

        // ── DELETE ──────────────────────────────────────────────────────
        Rule::delete_stmt => {
            walk_delete_stmt(pair)?
        }

        // ── START ───────────────────────────────────────────────────────
        Rule::start_stmt => {
            // START positions file — simplified as empty
            StmtKind::Empty
        }

        // ── SORT ────────────────────────────────────────────────────────
        Rule::sort_stmt => {
            // Sort is complex — simplified as empty for now
            StmtKind::Empty
        }

        // ── MERGE ───────────────────────────────────────────────────────
        Rule::merge_stmt => {
            StmtKind::Empty
        }

        // ── SEARCH ──────────────────────────────────────────────────────
        Rule::search_stmt => {
            walk_search_stmt(pair)?
        }

        // ── COPY ────────────────────────────────────────────────────────
        Rule::copy_stmt => {
            // Preprocessor directive — no runtime effect
            StmtKind::Empty
        }

        // ── INVOKE (OO COBOL) ──────────────────────────────────────────
        Rule::invoke_stmt => {
            walk_invoke_stmt(pair)?
        }

        // ── VALIDATE ────────────────────────────────────────────────────
        Rule::validate_stmt => {
            StmtKind::Empty
        }

        // ── FREE ────────────────────────────────────────────────────────
        Rule::free_stmt => {
            let parts = inner_nokw(pair);
            let name = parts.into_iter()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            StmtKind::Assign {
                targets: vec![Expression::ident(&name)],
                value: Expression::null(),
            }
        }

        // ── ALLOCATE ────────────────────────────────────────────────────
        Rule::allocate_stmt => {
            let parts = inner_nokw(pair);
            let names: Vec<String> = parts.into_iter()
                .filter(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
                .collect();
            let target = names.first().cloned().unwrap_or_default();
            StmtKind::Assign {
                targets: vec![Expression::ident(&target)],
                value: Expression::new(ExprKind::Object(Vec::new())),
            }
        }

        // ── TYPEDEF ─────────────────────────────────────────────────────
        Rule::typedef_stmt => {
            StmtKind::Empty
        }

        // ── EXIT ────────────────────────────────────────────────────────
        Rule::exit_stmt => {
            walk_exit_stmt(pair)?
        }

        // ── WAIT ────────────────────────────────────────────────────────
        Rule::wait_stmt => {
            let parts = inner_nokw(pair);
            let name = parts.into_iter()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            StmtKind::Expr(Expression::new(ExprKind::Await(
                Box::new(Expression::ident(&name)),
            )))
        }

        // ── RUN UNIT ────────────────────────────────────────────────────
        Rule::run_unit_stmt => {
            walk_run_unit_stmt(pair)?
        }

        // ── LOCK ────────────────────────────────────────────────────────
        Rule::lock_stmt => {
            let parts = inner_nokw(pair);
            let name = parts.into_iter()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__lock")),
                args: vec![Argument::positional(Expression::ident(&name))],
                optional: false,
            }))
        }

        // ── UNLOCK ──────────────────────────────────────────────────────
        Rule::unlock_stmt => {
            let parts = inner_nokw(pair);
            let name = parts.into_iter()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__unlock")),
                args: vec![Argument::positional(Expression::ident(&name))],
                optional: false,
            }))
        }

        // ── YIELD ───────────────────────────────────────────────────────
        Rule::yield_stmt => {
            StmtKind::Expr(Expression::new(ExprKind::Yield(None)))
        }

        // ── SUSPEND ─────────────────────────────────────────────────────
        Rule::suspend_stmt => {
            StmtKind::Expr(Expression::new(ExprKind::Yield(None)))
        }

        // ── EXEC SQL ────────────────────────────────────────────────────
        Rule::exec_sql_stmt => {
            let raw = extract_exec_body(&pair);
            StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__exec_sql")),
                args: vec![Argument::positional(Expression::string(&raw))],
                optional: false,
            }))
        }

        // ── EXEC CICS ───────────────────────────────────────────────────
        Rule::exec_cics_stmt => {
            let raw = extract_exec_body(&pair);
            StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__exec_cics")),
                args: vec![Argument::positional(Expression::string(&raw))],
                optional: false,
            }))
        }

        // ── EXEC DLI ────────────────────────────────────────────────────
        Rule::exec_dli_stmt => {
            let raw = extract_exec_body(&pair);
            StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__exec_dli")),
                args: vec![Argument::positional(Expression::string(&raw))],
                optional: false,
            }))
        }

        // ── Nested program ──────────────────────────────────────────────
        Rule::nested_program_stmt => {
            // Compile inline — walk its procedure division
            let mut nested_body = Vec::new();
            for child in pair.into_inner() {
                match child.as_rule() {
                    Rule::procedure_division => {
                        walk_procedure_division(child, &mut nested_body)?;
                    }
                    Rule::data_division => {
                        walk_data_division(child, &mut nested_body)?;
                    }
                    _ => {}
                }
            }
            StmtKind::Block(nested_body)
        }

        // ── statement_list (transparent wrapper) ────────────────────────
        Rule::statement_list => {
            let mut stmts = Vec::new();
            walk_statement_list(pair, &mut stmts)?;
            if stmts.len() == 1 {
                return Ok(Some(stmts.remove(0)));
            }
            StmtKind::Block(stmts)
        }

        // Skip periods and EOI
        Rule::period | Rule::EOI => return Ok(None),

        other => {
            if is_kw(other) {
                return Ok(None);
            }
            return Err(format!("COBOL walker: unhandled statement rule {:?}", other));
        }
    };

    Ok(Some(Statement::with_span(kind, span)))
}

// ════════════════════════════════════════════════════════════════════════════
// Statement walkers
// ════════════════════════════════════════════════════════════════════════════

// ── ACCEPT ──────────────────────────────────────────────────────────────────

fn walk_accept_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    let mut var_name = String::new();
    let mut source: Option<String> = None;

    for child in &children {
        match child.as_rule() {
            Rule::ident_name => {
                if var_name.is_empty() {
                    var_name = child.as_str().to_string();
                }
            }
            Rule::accept_source => {
                // Determine the accept source
                for inner in child.clone().into_inner() {
                    let s = match inner.as_rule() {
                        Rule::kw_date => "DATE",
                        Rule::kw_time => "TIME",
                        Rule::kw_day => "DAY",
                        Rule::kw_day_of_week => "DAY-OF-WEEK",
                        Rule::kw_command_line => "COMMAND-LINE",
                        Rule::kw_console => "CONSOLE",
                        _ => continue,
                    };
                    source = Some(s.to_string());
                }
            }
            _ => {}
        }
    }

    let call_callee = match source.as_deref() {
        Some("DATE") | Some("DAY") | Some("DAY-OF-WEEK") => "__accept_date",
        Some("TIME") => "__accept_time",
        Some("COMMAND-LINE") => "__accept_command_line",
        _ => "readline",
    };

    // ACCEPT var → var = readline()
    Ok(StmtKind::Assign {
        targets: vec![Expression::ident(&var_name)],
        value: Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(call_callee)),
            args: Vec::new(),
            optional: false,
        }),
    })
}

// ── MOVE ────────────────────────────────────────────────────────────────────

fn walk_move_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Check for MOVE CORRESPONDING
    let has_corr = children.iter().any(|c|
        matches!(c.as_rule(), Rule::kw_corresponding | Rule::kw_corr)
    );

    if has_corr {
        // MOVE CORRESPONDING src TO dst
        let idents: Vec<String> = children.iter()
            .filter(|c| c.as_rule() == Rule::ident_name)
            .map(|c| c.as_str().to_string())
            .collect();
        let src = idents.first().cloned().unwrap_or_default();
        let dst = idents.get(1).cloned().unwrap_or_default();

        // Emit as a call to __move_corresponding(src, dst)
        Ok(StmtKind::Expr(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__move_corresponding")),
            args: vec![
                Argument::positional(Expression::ident(&src)),
                Argument::positional(Expression::ident(&dst)),
            ],
            optional: false,
        })))
    } else {
        // MOVE expression TO ident+
        let mut src_expr: Option<Expression> = None;
        let mut targets = Vec::new();
        let mut after_to = false;

        for child in children {
            if child.as_rule() == Rule::kw_to {
                after_to = true;
                continue;
            }
            if is_kw(child.as_rule()) { continue; }

            if !after_to {
                if child.as_rule() == Rule::expression {
                    src_expr = Some(walk_expression(child)?);
                }
            } else if child.as_rule() == Rule::ident_name {
                targets.push(Expression::ident(child.as_str()));
            }
        }

        let value = src_expr.ok_or("MOVE missing source expression")?;
        Ok(StmtKind::Assign { targets, value })
    }
}

// ── ADD ─────────────────────────────────────────────────────────────────────

fn walk_add_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Check for CORRESPONDING
    let has_corr = children.iter().any(|c|
        matches!(c.as_rule(), Rule::kw_corresponding | Rule::kw_corr)
    );

    if has_corr {
        let idents: Vec<String> = children.iter()
            .filter(|c| c.as_rule() == Rule::ident_name)
            .map(|c| c.as_str().to_string())
            .collect();
        let src = idents.first().cloned().unwrap_or_default();
        let dst = idents.get(1).cloned().unwrap_or_default();
        return Ok(StmtKind::CompoundAssign {
            target: Expression::ident(&dst),
            op: CompoundOp::Add,
            value: Expression::ident(&src),
        });
    }

    // Collect expressions before TO, identifiers after TO
    let mut exprs: Vec<Expression> = Vec::new();
    let mut giving_name: Option<String> = None;
    let mut to_name: Option<String> = None;
    let mut in_giving = false;
    let mut in_to = false;

    for child in &children {
        match child.as_rule() {
            Rule::kw_to => { in_to = true; in_giving = false; continue; }
            Rule::kw_giving => { in_giving = true; in_to = false; continue; }
            _ => {}
        }
        if is_kw(child.as_rule()) || matches!(child.as_rule(), Rule::size_error_clause | Rule::rounded_clause) {
            continue;
        }

        if in_giving {
            if child.as_rule() == Rule::ident_name {
                giving_name = Some(child.as_str().to_string());
            }
        } else if in_to {
            if child.as_rule() == Rule::ident_name {
                to_name = Some(child.as_str().to_string());
            } else if child.as_rule() == Rule::giving_clause {
                // giving_clause nested inside
                for gc in child.clone().into_inner() {
                    if gc.as_rule() == Rule::ident_name {
                        giving_name = Some(gc.as_str().to_string());
                    }
                }
            }
        } else if child.as_rule() == Rule::expression {
            exprs.push(walk_expression(child.clone())?);
        } else if child.as_rule() == Rule::giving_clause {
            for gc in child.clone().into_inner() {
                if gc.as_rule() == Rule::ident_name {
                    giving_name = Some(gc.as_str().to_string());
                }
            }
        }
    }

    // Build the sum expression
    let sum_expr = build_sum_expr(&exprs);

    if let Some(giving) = giving_name {
        // ADD a b GIVING c → c = a + b (+ to if present)
        let total = if let Some(ref to) = to_name {
            binary(BinOp::Add, sum_expr, Expression::ident(to))
        } else {
            sum_expr
        };
        Ok(StmtKind::Assign {
            targets: vec![Expression::ident(&giving)],
            value: total,
        })
    } else if let Some(to) = to_name {
        // ADD a TO b → b += a
        Ok(StmtKind::CompoundAssign {
            target: Expression::ident(&to),
            op: CompoundOp::Add,
            value: sum_expr,
        })
    } else {
        // Fallback: ADD a TO first expr
        Ok(StmtKind::Empty)
    }
}

// ── SUBTRACT ────────────────────────────────────────────────────────────────

fn walk_subtract_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let has_corr = children.iter().any(|c|
        matches!(c.as_rule(), Rule::kw_corresponding | Rule::kw_corr)
    );

    if has_corr {
        let idents: Vec<String> = children.iter()
            .filter(|c| c.as_rule() == Rule::ident_name)
            .map(|c| c.as_str().to_string())
            .collect();
        let src = idents.first().cloned().unwrap_or_default();
        let dst = idents.get(1).cloned().unwrap_or_default();
        return Ok(StmtKind::CompoundAssign {
            target: Expression::ident(&dst),
            op: CompoundOp::Sub,
            value: Expression::ident(&src),
        });
    }

    let mut src_expr: Option<Expression> = None;
    let mut from_name: Option<String> = None;
    let mut giving_name: Option<String> = None;
    let mut in_from = false;

    for child in &children {
        match child.as_rule() {
            Rule::kw_from => { in_from = true; continue; }
            _ => {}
        }
        if is_kw(child.as_rule()) || matches!(child.as_rule(), Rule::size_error_clause | Rule::rounded_clause) {
            continue;
        }
        if child.as_rule() == Rule::giving_clause {
            for gc in child.clone().into_inner() {
                if gc.as_rule() == Rule::ident_name {
                    giving_name = Some(gc.as_str().to_string());
                }
            }
            continue;
        }

        if in_from {
            if child.as_rule() == Rule::ident_name {
                from_name = Some(child.as_str().to_string());
            }
        } else if child.as_rule() == Rule::expression {
            src_expr = Some(walk_expression(child.clone())?);
        }
    }

    let src = src_expr.unwrap_or(Expression::int(0));

    if let Some(giving) = giving_name {
        let from_expr = from_name.map(|n| Expression::ident(&n)).unwrap_or(Expression::int(0));
        Ok(StmtKind::Assign {
            targets: vec![Expression::ident(&giving)],
            value: binary(BinOp::Sub, from_expr, src),
        })
    } else if let Some(from) = from_name {
        Ok(StmtKind::CompoundAssign {
            target: Expression::ident(&from),
            op: CompoundOp::Sub,
            value: src,
        })
    } else {
        Ok(StmtKind::Empty)
    }
}

// ── MULTIPLY ────────────────────────────────────────────────────────────────

fn walk_multiply_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut src_expr: Option<Expression> = None;
    let mut by_name: Option<String> = None;
    let mut giving_name: Option<String> = None;
    let mut in_by = false;

    for child in &children {
        match child.as_rule() {
            Rule::kw_by => { in_by = true; continue; }
            _ => {}
        }
        if is_kw(child.as_rule()) || matches!(child.as_rule(), Rule::size_error_clause | Rule::rounded_clause) {
            continue;
        }
        if child.as_rule() == Rule::giving_clause {
            for gc in child.clone().into_inner() {
                if gc.as_rule() == Rule::ident_name {
                    giving_name = Some(gc.as_str().to_string());
                }
            }
            continue;
        }

        if in_by {
            if child.as_rule() == Rule::ident_name {
                by_name = Some(child.as_str().to_string());
            }
        } else if child.as_rule() == Rule::expression {
            src_expr = Some(walk_expression(child.clone())?);
        }
    }

    let src = src_expr.unwrap_or(Expression::int(1));

    if let Some(giving) = giving_name {
        let by_expr = by_name.map(|n| Expression::ident(&n)).unwrap_or(Expression::int(1));
        Ok(StmtKind::Assign {
            targets: vec![Expression::ident(&giving)],
            value: binary(BinOp::Mul, src, by_expr),
        })
    } else if let Some(by) = by_name {
        Ok(StmtKind::CompoundAssign {
            target: Expression::ident(&by),
            op: CompoundOp::Mul,
            value: src,
        })
    } else {
        Ok(StmtKind::Empty)
    }
}

// ── DIVIDE ──────────────────────────────────────────────────────────────────

fn walk_divide_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut exprs: Vec<Expression> = Vec::new();
    let mut giving_name: Option<String> = None;
    let mut remainder_name: Option<String> = None;
    let mut is_by = false;
    let mut is_into = false;

    for child in &children {
        match child.as_rule() {
            Rule::kw_by => { is_by = true; is_into = false; continue; }
            Rule::kw_into => { is_into = true; is_by = false; continue; }
            _ => {}
        }
        if is_kw(child.as_rule()) || matches!(child.as_rule(), Rule::size_error_clause | Rule::rounded_clause) {
            continue;
        }
        if child.as_rule() == Rule::remainder_clause {
            for rc in child.clone().into_inner() {
                if rc.as_rule() == Rule::ident_name {
                    remainder_name = Some(rc.as_str().to_string());
                }
            }
            continue;
        }

        if child.as_rule() == Rule::expression {
            exprs.push(walk_expression(child.clone())?);
        } else if child.as_rule() == Rule::ident_name && (is_by || is_into) {
            // After GIVING keyword
            if giving_name.is_none() {
                giving_name = Some(child.as_str().to_string());
            }
        }
    }

    // DIVIDE a BY b GIVING c / DIVIDE a INTO b GIVING c
    let (dividend, divisor) = if exprs.len() >= 2 {
        if is_into {
            // DIVIDE a INTO b → b / a
            (exprs[1].clone(), exprs[0].clone())
        } else {
            // DIVIDE a BY b → a / b
            (exprs[0].clone(), exprs[1].clone())
        }
    } else if exprs.len() == 1 {
        (exprs[0].clone(), Expression::int(1))
    } else {
        (Expression::int(0), Expression::int(1))
    };

    let target_name = giving_name.unwrap_or_default();

    if let Some(rem_name) = remainder_name {
        // Two assigns: c = a / b, r = a % b
        // Wrap in a block
        let div_assign = Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident(&target_name)],
            value: binary(BinOp::IDiv, dividend.clone(), divisor.clone()),
        });
        let rem_assign = Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident(&rem_name)],
            value: binary(BinOp::Mod, dividend, divisor),
        });
        Ok(StmtKind::Block(vec![div_assign, rem_assign]))
    } else {
        Ok(StmtKind::Assign {
            targets: vec![Expression::ident(&target_name)],
            value: binary(BinOp::Div, dividend, divisor),
        })
    }
}

// ── COMPUTE ─────────────────────────────────────────────────────────────────

fn walk_compute_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let parts = inner_nokw(pair);
    let mut target = String::new();
    let mut expr: Option<Expression> = None;

    for p in parts {
        match p.as_rule() {
            Rule::ident_name => {
                if target.is_empty() {
                    target = p.as_str().to_string();
                }
            }
            Rule::expression => {
                expr = Some(walk_expression(p)?);
            }
            _ => {}
        }
    }

    Ok(StmtKind::Assign {
        targets: vec![Expression::ident(&target)],
        value: expr.unwrap_or(Expression::int(0)),
    })
}

// ── IF ──────────────────────────────────────────────────────────────────────

fn walk_if_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut cond: Option<Expression> = None;
    let mut then_body = Vec::new();
    let mut else_body: Option<Vec<Statement>> = None;

    for child in children {
        match child.as_rule() {
            Rule::condition => {
                if cond.is_none() {
                    cond = Some(walk_condition(child)?);
                }
            }
            Rule::statement_list => {
                if cond.is_some() && then_body.is_empty() {
                    walk_statement_list(child, &mut then_body)?;
                }
            }
            Rule::else_clause => {
                let mut body = Vec::new();
                for ec in child.into_inner() {
                    if ec.as_rule() == Rule::statement_list {
                        walk_statement_list(ec, &mut body)?;
                    }
                }
                else_body = Some(body);
            }
            _ => {}
        }
    }

    Ok(StmtKind::If {
        cond: cond.unwrap_or(Expression::bool(true)),
        then_body,
        elifs: Vec::new(),
        else_body,
    })
}

// ── EVALUATE ────────────────────────────────────────────────────────────────

fn walk_evaluate_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // First expression is the subject
    let mut subject: Option<Expression> = None;
    let mut cases: Vec<SwitchCase> = Vec::new();
    let mut default: Option<Vec<Statement>> = None;

    for child in children {
        match child.as_rule() {
            Rule::expression => {
                if subject.is_none() {
                    subject = Some(walk_expression(child)?);
                }
            }
            Rule::when_clause => {
                let case = walk_when_clause(child)?;
                cases.push(case);
            }
            Rule::when_other_clause => {
                let mut body = Vec::new();
                for wc in child.into_inner() {
                    if wc.as_rule() == Rule::statement_list {
                        walk_statement_list(wc, &mut body)?;
                    }
                }
                default = Some(body);
            }
            _ => {}
        }
    }

    let subject_expr = subject.unwrap_or(Expression::bool(true));

    // If EVALUATE TRUE, convert to if/elif chain
    let is_eval_true = matches!(&subject_expr.kind, ExprKind::Lit(Literal::Bool(true)));

    if is_eval_true {
        // Convert to If chain
        return convert_evaluate_true_to_if(cases, default);
    }

    Ok(StmtKind::Switch {
        expr: subject_expr,
        cases,
        default,
    })
}

fn walk_when_clause(pair: Pair<Rule>) -> Result<SwitchCase, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    let mut conditions = Vec::new();
    let mut body = Vec::new();

    for child in children {
        match child.as_rule() {
            Rule::when_value => {
                let val = walk_when_value(child)?;
                if let Some(cond) = val {
                    conditions.push(cond);
                }
            }
            Rule::statement_list => {
                walk_statement_list(child, &mut body)?;
            }
            _ => {}
        }
    }

    Ok(SwitchCase { conditions, body })
}

fn walk_when_value(pair: Pair<Rule>) -> Result<Option<CaseCondition>, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    for child in &children {
        match child.as_rule() {
            Rule::kw_any => return Ok(None), // ANY matches everything
            Rule::kw_true => return Ok(Some(CaseCondition::Value(Expression::bool(true)))),
            Rule::kw_false => return Ok(Some(CaseCondition::Value(Expression::bool(false)))),
            _ => {}
        }
    }

    // Check for THRU range
    let exprs: Vec<Expression> = children.iter()
        .filter(|c| c.as_rule() == Rule::expression)
        .map(|c| walk_expression(c.clone()))
        .collect::<Result<Vec<_>, _>>()?;

    if exprs.len() >= 2 {
        Ok(Some(CaseCondition::Range {
            from: exprs[0].clone(),
            to: exprs[1].clone(),
        }))
    } else if let Some(expr) = exprs.into_iter().next() {
        Ok(Some(CaseCondition::Value(expr)))
    } else {
        Ok(None)
    }
}

/// Convert EVALUATE TRUE with WHEN conditions to if/elif chain.
fn convert_evaluate_true_to_if(
    cases: Vec<SwitchCase>,
    default: Option<Vec<Statement>>,
) -> Result<StmtKind, String> {
    if cases.is_empty() {
        return Ok(StmtKind::Block(default.unwrap_or_default()));
    }

    let mut iter = cases.into_iter();
    let first = iter.next().unwrap();

    let cond = case_conditions_to_expr(&first.conditions);
    let then_body = first.body;

    let mut elifs: Vec<(Expression, Vec<Statement>)> = Vec::new();
    for case in iter {
        let elif_cond = case_conditions_to_expr(&case.conditions);
        elifs.push((elif_cond, case.body));
    }

    Ok(StmtKind::If {
        cond,
        then_body,
        elifs,
        else_body: default,
    })
}

/// Extract the condition expression from WHEN clause conditions for EVALUATE TRUE.
fn case_conditions_to_expr(conditions: &[CaseCondition]) -> Expression {
    if conditions.is_empty() {
        return Expression::bool(true);
    }
    match &conditions[0] {
        CaseCondition::Value(expr) => expr.clone(),
        CaseCondition::Range { from, to } => {
            // value >= from AND value <= to — but for EVALUATE TRUE this is a condition
            binary(BinOp::And,
                binary(BinOp::GtEq, from.clone(), from.clone()),
                binary(BinOp::LtEq, to.clone(), to.clone()),
            )
        }
        CaseCondition::Comparison { op: _, expr } => expr.clone(),
    }
}

// ── PERFORM ─────────────────────────────────────────────────────────────────

fn walk_perform_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let inner = pair.into_inner().next().ok_or("empty PERFORM")?;

    match inner.as_rule() {
        Rule::perform_varying => walk_perform_varying(inner),
        Rule::perform_until => walk_perform_until(inner),
        Rule::perform_times => walk_perform_times(inner),
        Rule::perform_async => {
            let parts = inner_nokw(inner);
            let name = parts.into_iter()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            Ok(StmtKind::Expr(Expression::new(ExprKind::Await(
                Box::new(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(&name)),
                    args: Vec::new(),
                    optional: false,
                })),
            ))))
        }
        Rule::perform_thru => {
            let parts = inner_nokw(inner);
            let names: Vec<String> = parts.into_iter()
                .filter(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
                .collect();
            // PERFORM para1 THRU para2 → call both paragraphs
            let mut stmts = Vec::new();
            for name in &names {
                stmts.push(Statement::new(StmtKind::Expr(
                    Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(name)),
                        args: Vec::new(),
                        optional: false,
                    }),
                )));
            }
            Ok(StmtKind::Block(stmts))
        }
        Rule::perform_paragraph => {
            let parts = inner_nokw(inner);
            let name = parts.into_iter()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            Ok(StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(&name)),
                args: Vec::new(),
                optional: false,
            })))
        }
        Rule::perform_inline => {
            let mut body = Vec::new();
            for child in inner.into_inner() {
                if child.as_rule() == Rule::statement_list {
                    walk_statement_list(child, &mut body)?;
                }
            }
            Ok(StmtKind::Block(body))
        }
        other => Err(format!("COBOL walker: unhandled perform variant {:?}", other)),
    }
}

fn walk_perform_varying(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut var_name = String::new();
    let mut from_expr: Option<Expression> = None;
    let mut by_expr: Option<Expression> = None;
    let mut until_cond: Option<Expression> = None;
    let mut body = Vec::new();

    let mut state = 0; // 0=pre-from, 1=from, 2=by, 3=until

    for child in children {
        match child.as_rule() {
            Rule::ident_name => {
                if var_name.is_empty() {
                    var_name = child.as_str().to_string();
                }
            }
            Rule::kw_from => { state = 1; continue; }
            Rule::kw_by => { state = 2; continue; }
            Rule::kw_until => { state = 3; continue; }
            Rule::expression => {
                match state {
                    1 => from_expr = Some(walk_expression(child)?),
                    2 => by_expr = Some(walk_expression(child)?),
                    _ => {}
                }
            }
            Rule::condition => {
                until_cond = Some(walk_condition(child)?);
            }
            Rule::statement_list => {
                walk_statement_list(child, &mut body)?;
            }
            _ => {}
        }
    }

    let from = from_expr.unwrap_or(Expression::int(0));
    let by = by_expr.unwrap_or(Expression::int(1));
    // COBOL UNTIL = loop while NOT condition
    let cond = negate_expr(until_cond.unwrap_or(Expression::bool(false)));

    // PERFORM VARYING var FROM start BY step UNTIL cond
    // → for (var = start; NOT cond; var += step) { body }
    let init = Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(&var_name)],
        value: from,
    });

    let update = Expression::new(ExprKind::Assign {
        target: Box::new(Expression::ident(&var_name)),
        value: Box::new(binary(BinOp::Add, Expression::ident(&var_name), by)),
    });

    Ok(StmtKind::For {
        init: Some(Box::new(init)),
        cond: Some(cond),
        update: Some(update),
        body,
    })
}

fn walk_perform_until(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut test_after = false;
    let mut until_cond: Option<Expression> = None;
    let mut body = Vec::new();

    for child in children {
        match child.as_rule() {
            Rule::test_clause => {
                // WITH TEST BEFORE|AFTER
                for tc in child.into_inner() {
                    if tc.as_rule() == Rule::kw_after {
                        test_after = true;
                    }
                }
            }
            Rule::condition => {
                until_cond = Some(walk_condition(child)?);
            }
            Rule::statement_list => {
                walk_statement_list(child, &mut body)?;
            }
            _ => {}
        }
    }

    // COBOL UNTIL = while NOT condition
    let cond = negate_expr(until_cond.unwrap_or(Expression::bool(false)));

    if test_after {
        // WITH TEST AFTER → do-while
        Ok(StmtKind::DoWhile {
            body,
            cond,
            until: false,
        })
    } else {
        Ok(StmtKind::While {
            cond,
            body,
            else_body: None,
        })
    }
}

fn walk_perform_times(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut count_expr: Option<Expression> = None;
    let mut body = Vec::new();

    for child in children {
        match child.as_rule() {
            Rule::expression => {
                if count_expr.is_none() {
                    count_expr = Some(walk_expression(child)?);
                }
            }
            Rule::statement_list => {
                walk_statement_list(child, &mut body)?;
            }
            _ => {}
        }
    }

    let n = count_expr.unwrap_or(Expression::int(0));
    let counter = "__i";

    let init = Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(counter)],
        value: Expression::int(0),
    });
    let cond = binary(BinOp::Lt, Expression::ident(counter), n);
    let update = Expression::new(ExprKind::Assign {
        target: Box::new(Expression::ident(counter)),
        value: Box::new(binary(BinOp::Add, Expression::ident(counter), Expression::int(1))),
    });

    Ok(StmtKind::For {
        init: Some(Box::new(init)),
        cond: Some(cond),
        update: Some(update),
        body,
    })
}

// ── STRING ──────────────────────────────────────────────────────────────────

fn walk_string_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut source_exprs: Vec<Expression> = Vec::new();
    let mut into_name = String::new();

    for child in children {
        match child.as_rule() {
            Rule::string_source => {
                // Each string source has an expression (the value to concatenate)
                for sc in child.into_inner() {
                    if sc.as_rule() == Rule::expression {
                        source_exprs.push(walk_expression(sc)?);
                        break; // Take just the value, skip DELIMITED BY
                    }
                }
            }
            Rule::ident_name => {
                // The INTO target
                into_name = child.as_str().to_string();
            }
            _ => {}
        }
    }

    // Build concatenation: source1 & source2 & ...
    let concat_expr = if source_exprs.is_empty() {
        Expression::string("")
    } else {
        let mut result = source_exprs.remove(0);
        for src in source_exprs {
            result = binary(BinOp::Concat, result, src);
        }
        result
    };

    Ok(StmtKind::Assign {
        targets: vec![Expression::ident(&into_name)],
        value: concat_expr,
    })
}

// ── UNSTRING ────────────────────────────────────────────────────────────────

fn walk_unstring_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut src_name = String::new();
    let mut target_names: Vec<String> = Vec::new();
    let mut delimiter: Option<Expression> = None;

    for child in &children {
        match child.as_rule() {
            Rule::ident_name => {
                if src_name.is_empty() {
                    src_name = child.as_str().to_string();
                }
            }
            Rule::string_literal => {
                if delimiter.is_none() {
                    delimiter = Some(walk_string_literal(child)?);
                }
            }
            Rule::unstring_target => {
                for ut in child.clone().into_inner() {
                    if ut.as_rule() == Rule::ident_name {
                        target_names.push(ut.as_str().to_string());
                        break;
                    }
                }
            }
            _ => {}
        }
    }

    // UNSTRING src DELIMITED BY delim INTO t1 t2 t3
    // → split(src, delim) then assign each element
    let split_call = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&src_name)),
            field: "split".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(
            delimiter.unwrap_or(Expression::string(" ")),
        )],
        optional: false,
    });

    let mut stmts = Vec::new();
    // temp = src.split(delim)
    stmts.push(Statement::new(StmtKind::VarDecl {
        kind: VarDeclKind::Dim,
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident("__split_result".to_string()),
            type_hint: None,
            init: Some(split_call),
            array_bounds: None,
            with_events: false,
        }],
    }));

    for (i, target) in target_names.iter().enumerate() {
        stmts.push(Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident(target)],
            value: Expression::new(ExprKind::Index {
                object: Box::new(Expression::ident("__split_result")),
                index: Box::new(Expression::int(i as i64)),
                null_safe: false,
            }),
        }));
    }

    Ok(StmtKind::Block(stmts))
}

// ── INSPECT ─────────────────────────────────────────────────────────────────

fn walk_inspect_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let inner = pair.into_inner().next().ok_or("empty INSPECT")?;

    match inner.as_rule() {
        Rule::inspect_tallying => walk_inspect_tallying(inner),
        Rule::inspect_replacing => walk_inspect_replacing(inner),
        Rule::inspect_converting => walk_inspect_converting(inner),
        other => Err(format!("COBOL walker: unhandled inspect variant {:?}", other)),
    }
}

fn walk_inspect_tallying(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let parts = inner_nokw(pair);
    let idents: Vec<String> = parts.iter()
        .filter(|p| p.as_rule() == Rule::ident_name)
        .map(|p| p.as_str().to_string())
        .collect();

    let var = idents.first().cloned().unwrap_or_default();
    let counter = idents.get(1).cloned().unwrap_or_default();

    // Find the target string/ident in tally phrases
    let mut target_expr = Expression::string(" ");
    for p in parts {
        if p.as_rule() == Rule::inspect_tally_phrase {
            for tp in p.into_inner() {
                if tp.as_rule() == Rule::string_literal {
                    target_expr = walk_string_literal(&tp)?;
                    break;
                }
                if tp.as_rule() == Rule::ident_name {
                    target_expr = Expression::ident(tp.as_str());
                    break;
                }
            }
        }
    }

    // counter = split(var, target).length - 1
    let split_call = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&var)),
            field: "split".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(target_expr)],
        optional: false,
    });

    let len_expr = Expression::new(ExprKind::Member {
        object: Box::new(split_call),
        field: "length".to_string(),
        null_safe: false,
    });

    Ok(StmtKind::Assign {
        targets: vec![Expression::ident(&counter)],
        value: binary(BinOp::Sub, len_expr, Expression::int(1)),
    })
}

fn walk_inspect_replacing(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let parts = inner_nokw(pair);
    let var = parts.iter()
        .find(|p| p.as_rule() == Rule::ident_name)
        .map(|p| p.as_str().to_string())
        .unwrap_or_default();

    let mut old_expr = Expression::string("");
    let mut new_expr = Expression::string("");

    for p in &parts {
        if p.as_rule() == Rule::inspect_replace_phrase {
            let mut found_by = false;
            for rp in p.clone().into_inner() {
                if rp.as_rule() == Rule::kw_by { found_by = true; continue; }
                if !found_by {
                    if rp.as_rule() == Rule::string_literal {
                        old_expr = walk_string_literal(&rp)?;
                    } else if rp.as_rule() == Rule::ident_name {
                        old_expr = Expression::ident(rp.as_str());
                    }
                } else {
                    if rp.as_rule() == Rule::string_literal {
                        new_expr = walk_string_literal(&rp)?;
                    } else if rp.as_rule() == Rule::ident_name {
                        new_expr = Expression::ident(rp.as_str());
                    }
                }
            }
        }
    }

    // var = var.replace(old, new)
    let replace_call = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&var)),
            field: "replace".to_string(),
            null_safe: false,
        })),
        args: vec![
            Argument::positional(old_expr),
            Argument::positional(new_expr),
        ],
        optional: false,
    });

    Ok(StmtKind::Assign {
        targets: vec![Expression::ident(&var)],
        value: replace_call,
    })
}

fn walk_inspect_converting(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let parts = inner_nokw(pair);
    let var = parts.iter()
        .find(|p| p.as_rule() == Rule::ident_name)
        .map(|p| p.as_str().to_string())
        .unwrap_or_default();

    // INSPECT var CONVERTING from TO to → character replacement
    let mut from_expr = Expression::string("");
    let mut to_expr = Expression::string("");
    let mut found_to = false;

    for p in &parts {
        match p.as_rule() {
            Rule::kw_to => { found_to = true; }
            Rule::string_literal => {
                if !found_to {
                    from_expr = walk_string_literal(p)?;
                } else {
                    to_expr = walk_string_literal(p)?;
                }
            }
            Rule::ident_name => {
                if p.as_str() != &var {
                    if !found_to {
                        from_expr = Expression::ident(p.as_str());
                    } else {
                        to_expr = Expression::ident(p.as_str());
                    }
                }
            }
            _ => {}
        }
    }

    let replace_call = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&var)),
            field: "replace".to_string(),
            null_safe: false,
        })),
        args: vec![
            Argument::positional(from_expr),
            Argument::positional(to_expr),
        ],
        optional: false,
    });

    Ok(StmtKind::Assign {
        targets: vec![Expression::ident(&var)],
        value: replace_call,
    })
}

// ── CALL ────────────────────────────────────────────────────────────────────

fn walk_call_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut callee_name = String::new();
    let mut args: Vec<Argument> = Vec::new();
    let mut returning_var: Option<String> = None;

    for child in children {
        match child.as_rule() {
            Rule::string_literal => {
                if callee_name.is_empty() {
                    // Strip quotes
                    let s = child.as_str();
                    callee_name = s[1..s.len()-1].to_string();
                }
            }
            Rule::ident_name => {
                if callee_name.is_empty() {
                    callee_name = child.as_str().to_string();
                }
            }
            Rule::call_arg => {
                for ca in child.into_inner() {
                    if ca.as_rule() == Rule::ident_name {
                        args.push(Argument::positional(Expression::ident(ca.as_str())));
                    }
                }
            }
            Rule::kw_returning => {}
            _ => {
                // Check for RETURNING clause after the keyword
                if child.as_rule() == Rule::ident_name && returning_var.is_none() {
                    returning_var = Some(child.as_str().to_string());
                }
            }
        }
    }

    let call_expr = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(&callee_name)),
        args,
        optional: false,
    });

    if let Some(ret_var) = returning_var {
        Ok(StmtKind::Assign {
            targets: vec![Expression::ident(&ret_var)],
            value: call_expr,
        })
    } else {
        Ok(StmtKind::Expr(call_expr))
    }
}

// ── INITIALIZE ──────────────────────────────────────────────────────────────

fn walk_initialize_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let parts = inner_nokw(pair);
    let name = parts.into_iter()
        .find(|p| p.as_rule() == Rule::ident_name)
        .map(|p| p.as_str().to_string())
        .unwrap_or_default();

    // INITIALIZE sets to default value (spaces for alpha, zeros for numeric)
    // Without knowing the PIC at walk time, default to empty string
    Ok(StmtKind::Assign {
        targets: vec![Expression::ident(&name)],
        value: Expression::string(""),
    })
}

// ── SET ─────────────────────────────────────────────────────────────────────

fn walk_set_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut target = String::new();
    let mut value: Option<Expression> = None;
    let mut is_up = false;
    let mut is_down = false;

    for child in &children {
        match child.as_rule() {
            Rule::ident_name => {
                if target.is_empty() {
                    target = child.as_str().to_string();
                }
            }
            Rule::kw_true => { value = Some(Expression::bool(true)); }
            Rule::kw_false => { value = Some(Expression::bool(false)); }
            Rule::kw_up => { is_up = true; }
            Rule::kw_down => { is_down = true; }
            Rule::expression => {
                value = Some(walk_expression(child.clone())?);
            }
            _ => {}
        }
    }

    if is_up {
        // SET var UP BY n → var += n
        Ok(StmtKind::CompoundAssign {
            target: Expression::ident(&target),
            op: CompoundOp::Add,
            value: value.unwrap_or(Expression::int(1)),
        })
    } else if is_down {
        // SET var DOWN BY n → var -= n
        Ok(StmtKind::CompoundAssign {
            target: Expression::ident(&target),
            op: CompoundOp::Sub,
            value: value.unwrap_or(Expression::int(1)),
        })
    } else {
        Ok(StmtKind::Assign {
            targets: vec![Expression::ident(&target)],
            value: value.unwrap_or(Expression::bool(true)),
        })
    }
}

// ── JSON ────────────────────────────────────────────────────────────────────

fn walk_json_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let has_generate = children.iter().any(|c| c.as_rule() == Rule::kw_generate);
    let has_parse = children.iter().any(|c| c.as_rule() == Rule::kw_parse);

    let idents: Vec<String> = children.iter()
        .filter(|c| c.as_rule() == Rule::ident_name)
        .map(|c| c.as_str().to_string())
        .collect();

    if has_generate {
        // JSON GENERATE dst FROM src → dst = json_encode(src)
        let dst = idents.first().cloned().unwrap_or_default();
        let src = idents.get(1).cloned().unwrap_or_default();
        Ok(StmtKind::Assign {
            targets: vec![Expression::ident(&dst)],
            value: Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("json_encode")),
                args: vec![Argument::positional(Expression::ident(&src))],
                optional: false,
            }),
        })
    } else if has_parse {
        // JSON PARSE src INTO dst → dst = json_decode(src)
        let src = idents.first().cloned().unwrap_or_default();
        let dst = idents.get(1).cloned().unwrap_or_default();
        Ok(StmtKind::Assign {
            targets: vec![Expression::ident(&dst)],
            value: Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("json_decode")),
                args: vec![Argument::positional(Expression::ident(&src))],
                optional: false,
            }),
        })
    } else {
        Ok(StmtKind::Empty)
    }
}

// ── File I/O ────────────────────────────────────────────────────────────────

fn walk_open_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    let mut stmts = Vec::new();
    let mut current_mode = String::new();

    for child in children {
        match child.as_rule() {
            Rule::file_open_mode => {
                // Determine mode string
                let mode_text = child.as_str().to_uppercase();
                current_mode = if mode_text.contains("INPUT") {
                    "r".to_string()
                } else if mode_text.contains("OUTPUT") {
                    "w".to_string()
                } else if mode_text.contains("EXTEND") {
                    "a".to_string()
                } else {
                    "rw".to_string()
                };
            }
            Rule::ident_name => {
                let file_name = child.as_str().to_string();
                stmts.push(Statement::new(StmtKind::Expr(
                    Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__file_open")),
                        args: vec![
                            Argument::positional(Expression::ident(&file_name)),
                            Argument::positional(Expression::string(&current_mode)),
                        ],
                        optional: false,
                    }),
                )));
            }
            _ => {}
        }
    }

    if stmts.len() == 1 {
        Ok(stmts.remove(0).kind)
    } else {
        Ok(StmtKind::Block(stmts))
    }
}

fn walk_close_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let parts = inner_nokw(pair);
    let names: Vec<String> = parts.into_iter()
        .filter(|p| p.as_rule() == Rule::ident_name)
        .map(|p| p.as_str().to_string())
        .collect();

    let mut stmts = Vec::new();
    for name in names {
        stmts.push(Statement::new(StmtKind::Expr(
            Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__file_close")),
                args: vec![Argument::positional(Expression::ident(&name))],
                optional: false,
            }),
        )));
    }

    if stmts.len() == 1 {
        Ok(stmts.remove(0).kind)
    } else {
        Ok(StmtKind::Block(stmts))
    }
}

fn walk_read_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut file_name = String::new();
    let mut into_var: Option<String> = None;
    let mut at_end_body = Vec::new();
    let mut not_at_end_body = Vec::new();

    for child in children {
        match child.as_rule() {
            Rule::ident_name => {
                if file_name.is_empty() {
                    file_name = child.as_str().to_string();
                } else if into_var.is_none() {
                    into_var = Some(child.as_str().to_string());
                }
            }
            Rule::at_end_clause => {
                for ac in child.into_inner() {
                    if ac.as_rule() == Rule::statement_list {
                        if at_end_body.is_empty() {
                            walk_statement_list(ac, &mut at_end_body)?;
                        } else {
                            walk_statement_list(ac, &mut not_at_end_body)?;
                        }
                    }
                }
            }
            _ => {}
        }
    }

    let read_call = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__file_read")),
        args: vec![Argument::positional(Expression::ident(&file_name))],
        optional: false,
    });

    if let Some(var) = into_var {
        Ok(StmtKind::Assign {
            targets: vec![Expression::ident(&var)],
            value: read_call,
        })
    } else {
        Ok(StmtKind::Expr(read_call))
    }
}

fn walk_write_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let parts = inner_nokw(pair);
    let mut record_name = String::new();
    let mut from_var: Option<String> = None;

    for p in parts {
        match p.as_rule() {
            Rule::ident_name => {
                if record_name.is_empty() {
                    record_name = p.as_str().to_string();
                } else if from_var.is_none() {
                    from_var = Some(p.as_str().to_string());
                }
            }
            _ => {}
        }
    }

    let data_expr = if let Some(var) = from_var {
        Expression::ident(&var)
    } else {
        Expression::ident(&record_name)
    };

    Ok(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__file_write")),
        args: vec![
            Argument::positional(Expression::ident(&record_name)),
            Argument::positional(data_expr),
        ],
        optional: false,
    })))
}

fn walk_rewrite_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let parts = inner_nokw(pair);
    let mut record_name = String::new();
    let mut from_var: Option<String> = None;

    for p in parts {
        if p.as_rule() == Rule::ident_name {
            if record_name.is_empty() {
                record_name = p.as_str().to_string();
            } else if from_var.is_none() {
                from_var = Some(p.as_str().to_string());
            }
        }
    }

    let data_expr = from_var.map(|v| Expression::ident(&v))
        .unwrap_or_else(|| Expression::ident(&record_name));

    Ok(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__file_rewrite")),
        args: vec![
            Argument::positional(Expression::ident(&record_name)),
            Argument::positional(data_expr),
        ],
        optional: false,
    })))
}

fn walk_delete_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let parts = inner_nokw(pair);
    let name = parts.into_iter()
        .find(|p| p.as_rule() == Rule::ident_name)
        .map(|p| p.as_str().to_string())
        .unwrap_or_default();

    Ok(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__file_delete")),
        args: vec![Argument::positional(Expression::ident(&name))],
        optional: false,
    })))
}

// ── SEARCH ──────────────────────────────────────────────────────────────────

fn walk_search_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut _table_name = String::new();
    let mut at_end_body = Vec::new();
    let mut when_clauses: Vec<(Expression, Vec<Statement>)> = Vec::new();

    for child in children {
        match child.as_rule() {
            Rule::ident_name => {
                if _table_name.is_empty() {
                    _table_name = child.as_str().to_string();
                }
            }
            Rule::condition => {
                let cond = walk_condition(child)?;
                when_clauses.push((cond, Vec::new()));
            }
            Rule::statement_list => {
                if let Some(last) = when_clauses.last_mut() {
                    walk_statement_list(child, &mut last.1)?;
                } else {
                    walk_statement_list(child, &mut at_end_body)?;
                }
            }
            _ => {}
        }
    }

    // SEARCH → if/elif chain of WHEN conditions
    if when_clauses.is_empty() {
        return Ok(StmtKind::Block(at_end_body));
    }

    let mut iter = when_clauses.into_iter();
    let (first_cond, first_body) = iter.next().unwrap();

    let elifs: Vec<(Expression, Vec<Statement>)> = iter.collect();
    let else_body = if at_end_body.is_empty() { None } else { Some(at_end_body) };

    Ok(StmtKind::If {
        cond: first_cond,
        then_body: first_body,
        elifs,
        else_body,
    })
}

// ── INVOKE (OO COBOL) ──────────────────────────────────────────────────────

fn walk_invoke_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let parts = inner_nokw(pair);
    let idents: Vec<String> = parts.iter()
        .filter(|p| p.as_rule() == Rule::ident_name)
        .map(|p| p.as_str().to_string())
        .collect();

    let obj = idents.first().cloned().unwrap_or_default();
    let method = idents.get(1).cloned().unwrap_or_default();
    let args: Vec<Argument> = idents.iter().skip(2)
        .map(|name| Argument::positional(Expression::ident(name)))
        .collect();

    Ok(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&obj)),
            field: method,
            null_safe: false,
        })),
        args,
        optional: false,
    })))
}

// ── EXIT ────────────────────────────────────────────────────────────────────

fn walk_exit_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let has_perform = children.iter().any(|c| c.as_rule() == Rule::kw_perform);
    let has_cycle = children.iter().any(|c| c.as_rule() == Rule::kw_cycle);
    let has_paragraph = children.iter().any(|c| c.as_rule() == Rule::kw_paragraph);
    let has_section = children.iter().any(|c| c.as_rule() == Rule::kw_section);
    let has_program = children.iter().any(|c| c.as_rule() == Rule::kw_program);
    let has_method = children.iter().any(|c| c.as_rule() == Rule::kw_method);

    if has_perform && has_cycle {
        // EXIT PERFORM CYCLE → continue
        Ok(StmtKind::Continue(ContinueTarget::Implicit))
    } else if has_perform {
        // EXIT PERFORM → break
        Ok(StmtKind::Break(BreakTarget::Implicit))
    } else if has_paragraph || has_section || has_program || has_method {
        // EXIT PARAGRAPH/SECTION/PROGRAM/METHOD → return
        Ok(StmtKind::Return(None))
    } else {
        // Bare EXIT → no-op
        Ok(StmtKind::Empty)
    }
}

// ── RUN UNIT ────────────────────────────────────────────────────────────────

fn walk_run_unit_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut name = String::new();
    let mut args: Vec<Argument> = Vec::new();
    let mut in_using = false;

    for child in children {
        match child.as_rule() {
            Rule::string_literal => {
                if name.is_empty() {
                    let s = child.as_str();
                    name = s[1..s.len()-1].to_string();
                }
            }
            Rule::ident_name => {
                if name.is_empty() {
                    name = child.as_str().to_string();
                } else if in_using {
                    args.push(Argument::positional(Expression::ident(child.as_str())));
                }
            }
            Rule::kw_using => { in_using = true; }
            _ => {}
        }
    }

    Ok(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(&name)),
        args,
        optional: false,
    })))
}

// ── Nested program ──────────────────────────────────────────────────────────

fn walk_nested_program(pair: Pair<Rule>, body: &mut Vec<Statement>) -> Result<(), String> {
    for child in pair.into_inner() {
        if child.as_rule() == Rule::nested_program_stmt {
            if let Some(stmt) = walk_statement(child)? {
                body.push(stmt);
            }
        }
    }
    Ok(())
}

// ── EXEC body extraction ────────────────────────────────────────────────────

fn extract_exec_body(pair: &Pair<Rule>) -> String {
    for child in pair.clone().into_inner() {
        if child.as_rule() == Rule::exec_body {
            return child.as_str().trim().to_string();
        }
    }
    String::new()
}

// ════════════════════════════════════════════════════════════════════════════
// OO COBOL: Classes and Interfaces
// ════════════════════════════════════════════════════════════════════════════

fn walk_class_id(pair: Pair<Rule>, body: &mut Vec<Statement>) -> Result<(), String> {
    let span = to_span(&pair);
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut name = String::new();
    let mut parents: Vec<String> = Vec::new();
    let mut interfaces: Vec<String> = Vec::new();
    let mut members: Vec<ClassMember> = Vec::new();

    for child in children {
        match child.as_rule() {
            Rule::ident_name => {
                if name.is_empty() {
                    name = child.as_str().to_string();
                }
            }
            Rule::inherits_clause => {
                for ic in child.into_inner() {
                    if ic.as_rule() == Rule::ident_name {
                        parents.push(ic.as_str().to_string());
                    }
                }
            }
            Rule::implements_clause => {
                for ic in child.into_inner() {
                    if ic.as_rule() == Rule::ident_name {
                        interfaces.push(ic.as_str().to_string());
                    }
                }
            }
            Rule::class_body => {
                walk_class_body(child, &mut members)?;
            }
            _ => {}
        }
    }

    body.push(Statement::with_span(
        StmtKind::ClassDecl {
            name,
            parents,
            interfaces,
            members,
            modifiers: ClassModifiers::default(),
        },
        span,
    ));
    Ok(())
}

fn walk_class_body(pair: Pair<Rule>, members: &mut Vec<ClassMember>) -> Result<(), String> {
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::data_division => {
                // Class-level data division → fields
                let mut field_stmts = Vec::new();
                walk_data_division(child, &mut field_stmts)?;
                for stmt in field_stmts {
                    if let StmtKind::VarDecl { declarations, .. } = &stmt.kind {
                        for decl in declarations {
                            if let BindingPattern::Ident(name) = &decl.pattern {
                                members.push(ClassMember::Field {
                                    name: name.clone(),
                                    type_hint: decl.type_hint.clone(),
                                    init: decl.init.clone(),
                                    modifiers: Modifiers::default(),
                                    with_events: false,
                                    array_bounds: None,
                                });
                            }
                        }
                    }
                }
            }
            Rule::method_def => {
                let method = walk_method_def(child)?;
                members.push(ClassMember::Method(Box::new(method)));
            }
            Rule::factory_paragraph => {
                // Factory methods → static methods
                for fc in child.into_inner() {
                    if fc.as_rule() == Rule::method_def {
                        let mut method = walk_method_def(fc)?;
                        if let StmtKind::FunctionDecl { ref mut modifiers, .. } = method.kind {
                            modifiers.is_static = true;
                        }
                        members.push(ClassMember::Method(Box::new(method)));
                    }
                }
            }
            Rule::object_paragraph => {
                for oc in child.into_inner() {
                    match oc.as_rule() {
                        Rule::data_division => {
                            let mut field_stmts = Vec::new();
                            walk_data_division(oc, &mut field_stmts)?;
                            for stmt in field_stmts {
                                if let StmtKind::VarDecl { declarations, .. } = &stmt.kind {
                                    for decl in declarations {
                                        if let BindingPattern::Ident(name) = &decl.pattern {
                                            members.push(ClassMember::Field {
                                                name: name.clone(),
                                                type_hint: decl.type_hint.clone(),
                                                init: decl.init.clone(),
                                                modifiers: Modifiers::default(),
                                                with_events: false,
                                                array_bounds: None,
                                            });
                                        }
                                    }
                                }
                            }
                        }
                        Rule::method_def => {
                            let method = walk_method_def(oc)?;
                            members.push(ClassMember::Method(Box::new(method)));
                        }
                        _ => {}
                    }
                }
            }
            Rule::procedure_division_class => {
                for pc in child.into_inner() {
                    if pc.as_rule() == Rule::method_def {
                        let method = walk_method_def(pc)?;
                        members.push(ClassMember::Method(Box::new(method)));
                    }
                }
            }
            _ => {}
        }
    }
    Ok(())
}

fn walk_method_def(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut name = String::new();
    let mut params: Vec<Param> = Vec::new();
    let mut return_type: Option<String> = None;
    let mut body = Vec::new();
    let mut is_override = false;
    let mut is_property_get = false;
    let mut is_property_set = false;

    for child in children {
        match child.as_rule() {
            Rule::ident_name => {
                if name.is_empty() {
                    name = child.as_str().to_string();
                }
            }
            Rule::method_modifiers => {
                for m in child.into_inner() {
                    match m.as_rule() {
                        Rule::kw_override => { is_override = true; }
                        Rule::property_modifier => {
                            for pm in m.into_inner() {
                                match pm.as_rule() {
                                    Rule::kw_get => { is_property_get = true; }
                                    Rule::kw_set => { is_property_set = true; }
                                    _ => {}
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            Rule::method_data_division => {
                // Method-local data → VarDecls in the body
                for md in child.into_inner() {
                    if md.as_rule() == Rule::data_item {
                        walk_data_item(md, &mut body)?;
                    }
                }
            }
            Rule::method_procedure_division => {
                for mp in child.into_inner() {
                    match mp.as_rule() {
                        Rule::using_clause => {
                            params = walk_using_params(mp);
                        }
                        Rule::returning_clause => {
                            for rc in mp.into_inner() {
                                if rc.as_rule() == Rule::ident_name {
                                    return_type = Some(rc.as_str().to_string());
                                }
                            }
                        }
                        Rule::statement_list => {
                            walk_statement_list(mp, &mut body)?;
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    let _ = is_property_get;
    let _ = is_property_set;

    let mut modifiers = Modifiers::default();
    modifiers.is_override = is_override;

    let is_sub = return_type.is_none();

    Ok(Statement::with_span(
        StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            body,
            modifiers,
            handles: Vec::new(),
            is_async: false,
            is_generator: false,
            is_sub,
        },
        span,
    ))
}

fn walk_using_params(pair: Pair<Rule>) -> Vec<Param> {
    let mut params = Vec::new();
    let mut pass_by = PassBy::Value;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::kw_reference => { pass_by = PassBy::Ref; }
            Rule::kw_content => { pass_by = PassBy::Const; }
            Rule::ident_name => {
                params.push(Param {
                    name: child.as_str().to_string(),
                    type_hint: None,
                    default: None,
                    pass_by,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                });
            }
            _ => {}
        }
    }
    params
}

fn walk_interface_id(pair: Pair<Rule>, body: &mut Vec<Statement>) -> Result<(), String> {
    let span = to_span(&pair);
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut name = String::new();
    let mut parents: Vec<String> = Vec::new();
    let mut members: Vec<InterfaceMember> = Vec::new();

    for child in children {
        match child.as_rule() {
            Rule::ident_name => {
                if name.is_empty() {
                    name = child.as_str().to_string();
                }
            }
            Rule::inherits_clause => {
                for ic in child.into_inner() {
                    if ic.as_rule() == Rule::ident_name {
                        parents.push(ic.as_str().to_string());
                    }
                }
            }
            Rule::interface_body => {
                for ib in child.into_inner() {
                    if ib.as_rule() == Rule::method_signature {
                        let sig = walk_method_signature(ib)?;
                        members.push(sig);
                    }
                }
            }
            _ => {}
        }
    }

    body.push(Statement::with_span(
        StmtKind::InterfaceDecl {
            name,
            parents,
            members,
        },
        span,
    ));
    Ok(())
}

fn walk_method_signature(pair: Pair<Rule>) -> Result<InterfaceMember, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    let mut name = String::new();
    let mut params: Vec<Param> = Vec::new();
    let mut return_type: Option<String> = None;

    for child in children {
        match child.as_rule() {
            Rule::ident_name => {
                if name.is_empty() {
                    name = child.as_str().to_string();
                }
            }
            Rule::using_clause => {
                params = walk_using_params(child);
            }
            Rule::returning_clause => {
                for rc in child.into_inner() {
                    if rc.as_rule() == Rule::ident_name {
                        return_type = Some(rc.as_str().to_string());
                    }
                }
            }
            _ => {}
        }
    }

    Ok(InterfaceMember::Method {
        name,
        params,
        return_type: return_type.clone(),
        is_sub: return_type.is_none(),
    })
}

// ════════════════════════════════════════════════════════════════════════════
// Conditions
// ════════════════════════════════════════════════════════════════════════════

fn walk_condition(pair: Pair<Rule>) -> Result<Expression, String> {
    let inner = pair.into_inner().next().ok_or("empty condition")?;
    walk_or_condition(inner)
}

fn walk_or_condition(pair: Pair<Rule>) -> Result<Expression, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    let mut result = walk_and_condition(children.first().ok_or("empty or_condition")?.clone())?;

    for i in (1..children.len()).step_by(1) {
        if children[i].as_rule() == Rule::and_condition {
            let right = walk_and_condition(children[i].clone())?;
            result = binary(BinOp::Or, result, right);
        }
    }
    Ok(result)
}

fn walk_and_condition(pair: Pair<Rule>) -> Result<Expression, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    let mut result = walk_not_condition(children.first().ok_or("empty and_condition")?.clone())?;

    for i in 1..children.len() {
        if children[i].as_rule() == Rule::not_condition {
            let right = walk_not_condition(children[i].clone())?;
            result = binary(BinOp::And, result, right);
        }
    }
    Ok(result)
}

fn walk_not_condition(pair: Pair<Rule>) -> Result<Expression, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let has_not = children.iter().any(|c| c.as_rule() == Rule::kw_not);
    let comparison = children.iter()
        .find(|c| c.as_rule() == Rule::comparison)
        .ok_or("not_condition without comparison")?;

    let expr = walk_comparison(comparison.clone())?;

    if has_not {
        Ok(Expression::new(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(expr),
        }))
    } else {
        Ok(expr)
    }
}

fn walk_comparison(pair: Pair<Rule>) -> Result<Expression, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Check for class condition test (IS NUMERIC, etc.)
    if let Some(class_test) = children.iter().find(|c| c.as_rule() == Rule::class_condition_test) {
        let expr = walk_expression(children.first().unwrap().clone())?;
        return walk_class_condition_test(class_test.clone(), expr);
    }

    // Check for sign condition test (IS POSITIVE, etc.)
    if let Some(sign_test) = children.iter().find(|c| c.as_rule() == Rule::sign_condition_test) {
        let expr = walk_expression(children.first().unwrap().clone())?;
        return walk_sign_condition_test(sign_test.clone(), expr);
    }

    // Check for comparison operator
    if let Some(comp_op) = children.iter().find(|c| c.as_rule() == Rule::comparison_op) {
        let exprs: Vec<Expression> = children.iter()
            .filter(|c| c.as_rule() == Rule::expression)
            .map(|c| walk_expression(c.clone()))
            .collect::<Result<Vec<_>, _>>()?;

        if exprs.len() >= 2 {
            let op = parse_comparison_op(comp_op.clone());
            return Ok(binary(op, exprs[0].clone(), exprs[1].clone()));
        }
    }

    // Single expression (truthiness test)
    if let Some(expr_pair) = children.iter().find(|c| c.as_rule() == Rule::expression) {
        return walk_expression(expr_pair.clone());
    }

    Ok(Expression::bool(true))
}

fn walk_class_condition_test(pair: Pair<Rule>, expr: Expression) -> Result<Expression, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    let is_negated = children.iter().any(|c| c.as_rule() == Rule::kw_not);

    let func_name = if children.iter().any(|c| c.as_rule() == Rule::kw_numeric) {
        "__is_numeric"
    } else if children.iter().any(|c| c.as_rule() == Rule::kw_alphabetic_lower) {
        "__is_alphabetic_lower"
    } else if children.iter().any(|c| c.as_rule() == Rule::kw_alphabetic_upper) {
        "__is_alphabetic_upper"
    } else {
        "__is_alphabetic"
    };

    let call = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(func_name)),
        args: vec![Argument::positional(expr)],
        optional: false,
    });

    if is_negated {
        Ok(negate_expr(call))
    } else {
        Ok(call)
    }
}

fn walk_sign_condition_test(pair: Pair<Rule>, expr: Expression) -> Result<Expression, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    let is_negated = children.iter().any(|c| c.as_rule() == Rule::kw_not);

    let result = if children.iter().any(|c| matches!(c.as_rule(), Rule::kw_positive)) {
        binary(BinOp::Gt, expr, Expression::int(0))
    } else if children.iter().any(|c| matches!(c.as_rule(), Rule::kw_negative)) {
        binary(BinOp::Lt, expr, Expression::int(0))
    } else {
        // ZERO/ZEROS/ZEROES
        binary(BinOp::Eq, expr, Expression::int(0))
    };

    if is_negated {
        Ok(negate_expr(result))
    } else {
        Ok(result)
    }
}

fn parse_comparison_op(pair: Pair<Rule>) -> BinOp {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    let text = children.iter().map(|c| c.as_str().to_uppercase()).collect::<Vec<_>>().join(" ");
    let is_negated = children.iter().any(|c| c.as_rule() == Rule::kw_not);

    // Check for symbolic operators first
    let raw_text = children.iter().map(|c| c.as_str()).collect::<String>();
    if raw_text.contains(">=") { return BinOp::GtEq; }
    if raw_text.contains("<=") { return BinOp::LtEq; }
    if raw_text.contains(">") { return if is_negated { BinOp::LtEq } else { BinOp::Gt }; }
    if raw_text.contains("<") { return if is_negated { BinOp::GtEq } else { BinOp::Lt }; }
    if raw_text.contains("=") && is_negated { return BinOp::NotEq; }
    if raw_text.contains("=") { return BinOp::Eq; }

    // Keyword operators
    if text.contains("EQUAL") {
        if is_negated { BinOp::NotEq } else { BinOp::Eq }
    } else if text.contains("GREATER") {
        if text.contains("EQUAL") {
            if is_negated { BinOp::Lt } else { BinOp::GtEq }
        } else {
            if is_negated { BinOp::LtEq } else { BinOp::Gt }
        }
    } else if text.contains("LESS") {
        if text.contains("EQUAL") {
            if is_negated { BinOp::Gt } else { BinOp::LtEq }
        } else {
            if is_negated { BinOp::GtEq } else { BinOp::Lt }
        }
    } else {
        BinOp::Eq
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Expressions
// ════════════════════════════════════════════════════════════════════════════

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner().next().ok_or("empty expression")?;
    walk_add_expr_with_span(inner, span)
}

fn walk_add_expr_with_span(pair: Pair<Rule>, _outer_span: Span) -> Result<Expression, String> {
    walk_add_expr(pair)
}

fn walk_add_expr(pair: Pair<Rule>) -> Result<Expression, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut result = walk_mul_expr(children.first().ok_or("empty add_expr")?.clone())?;

    let mut i = 1;
    while i < children.len() {
        if children[i].as_rule() == Rule::add_op {
            let op_text = children[i].as_str();
            let op = if op_text == "+" { BinOp::Add } else { BinOp::Sub };
            i += 1;
            if i < children.len() {
                let right = walk_mul_expr(children[i].clone())?;
                result = binary(op, result, right);
            }
        }
        i += 1;
    }
    Ok(result)
}

fn walk_mul_expr(pair: Pair<Rule>) -> Result<Expression, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut result = walk_power_expr(children.first().ok_or("empty mul_expr")?.clone())?;

    let mut i = 1;
    while i < children.len() {
        if children[i].as_rule() == Rule::mul_op {
            let op_text = children[i].as_str();
            let op = if op_text.starts_with('*') { BinOp::Mul } else { BinOp::Div };
            i += 1;
            if i < children.len() {
                let right = walk_power_expr(children[i].clone())?;
                result = binary(op, result, right);
            }
        }
        i += 1;
    }
    Ok(result)
}

fn walk_power_expr(pair: Pair<Rule>) -> Result<Expression, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let base = walk_unary_expr(children.first().ok_or("empty power_expr")?.clone())?;

    if children.len() > 1 {
        let exp = walk_unary_expr(children.last().unwrap().clone())?;
        Ok(binary(BinOp::Pow, base, exp))
    } else {
        Ok(base)
    }
}

fn walk_unary_expr(pair: Pair<Rule>) -> Result<Expression, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    if children.len() == 2 {
        let op_text = children[0].as_str();
        let expr = walk_atom(children[1].clone())?;
        if op_text == "-" {
            return Ok(Expression::new(ExprKind::Unary {
                op: UnaryOp::Neg,
                expr: Box::new(expr),
            }));
        }
        // + is a no-op
        return Ok(expr);
    }

    if let Some(atom) = children.first() {
        return walk_atom(atom.clone());
    }

    Ok(Expression::int(0))
}

fn walk_atom(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);

    match pair.as_rule() {
        Rule::atom => {
            let inner = pair.into_inner().next().ok_or("empty atom")?;
            walk_atom(inner)
        }

        Rule::function_call => {
            walk_function_call(pair)
        }

        Rule::new_expr => {
            walk_new_expr(pair)
        }

        Rule::paren_expr => {
            let inner = pair.into_inner().next().ok_or("empty paren_expr")?;
            walk_expression(inner)
        }

        Rule::literal => {
            walk_literal(pair)
        }

        Rule::qualified_ident => {
            walk_qualified_ident(pair)
        }

        // For cases where expression/add_expr/etc. show up directly
        Rule::expression | Rule::add_expr => {
            if pair.as_rule() == Rule::expression {
                walk_expression(pair)
            } else {
                walk_add_expr(pair)
            }
        }

        Rule::mul_expr => walk_mul_expr(pair),
        Rule::power_expr => walk_power_expr(pair),
        Rule::unary_expr => walk_unary_expr(pair),

        other => {
            // Fallback: treat as an identifier if it looks like one
            let text = pair.as_str();
            if !text.is_empty() && text.chars().next().map_or(false, |c| c.is_alphabetic()) {
                Ok(Expression::with_span(ExprKind::Ident(text.to_string()), span))
            } else {
                Err(format!("COBOL walker: unhandled atom rule {:?}", other))
            }
        }
    }
}

// ── Function calls ──────────────────────────────────────────────────────────

fn walk_function_call(pair: Pair<Rule>) -> Result<Expression, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut func_name = String::new();
    let mut args: Vec<Argument> = Vec::new();

    for child in children {
        match child.as_rule() {
            Rule::function_name => {
                func_name = child.as_str().to_uppercase();
            }
            Rule::arg_list => {
                for ac in child.into_inner() {
                    if ac.as_rule() == Rule::expression {
                        args.push(Argument::positional(walk_expression(ac)?));
                    }
                }
            }
            Rule::atom => {
                // FUNCTION name atom atom ... (alternate syntax without parens)
                args.push(Argument::positional(walk_atom(child)?));
            }
            _ => {}
        }
    }

    Ok(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(&func_name)),
        args,
        optional: false,
    }))
}

// ── NEW expression (OO COBOL) ──────────────────────────────────────────────

fn walk_new_expr(pair: Pair<Rule>) -> Result<Expression, String> {
    let parts = inner_nokw(pair);
    let mut class_name = String::new();
    let mut args: Vec<Argument> = Vec::new();

    for p in parts {
        match p.as_rule() {
            Rule::ident_name => {
                if class_name.is_empty() {
                    class_name = p.as_str().to_string();
                }
            }
            Rule::arg_list => {
                for ac in p.into_inner() {
                    if ac.as_rule() == Rule::expression {
                        args.push(Argument::positional(walk_expression(ac)?));
                    }
                }
            }
            _ => {}
        }
    }

    Ok(Expression::new(ExprKind::New {
        class: Box::new(Expression::ident(&class_name)),
        args,
    }))
}

// ── Qualified identifiers ──────────────────────────────────────────────────

fn walk_qualified_ident(pair: Pair<Rule>) -> Result<Expression, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Check for SELF
    if children.iter().any(|c| c.as_rule() == Rule::kw_self) {
        return Ok(Expression::new(ExprKind::This));
    }

    let mut name = String::new();
    let mut subscript: Option<Pair<Rule>> = None;
    let mut qualification: Option<String> = None;

    for child in &children {
        match child.as_rule() {
            Rule::ident_name => {
                if name.is_empty() {
                    name = child.as_str().to_string();
                }
            }
            Rule::subscript_or_refmod => {
                subscript = Some(child.clone());
            }
            Rule::qualification => {
                for q in child.clone().into_inner() {
                    if q.as_rule() == Rule::ident_name {
                        qualification = Some(q.as_str().to_string());
                    }
                }
            }
            _ => {}
        }
    }

    let mut expr = Expression::ident(&name);

    // Handle subscript or reference modification
    if let Some(sub_pair) = subscript {
        let sub_children: Vec<Pair<Rule>> = sub_pair.into_inner().collect();

        // Check if it's reference modification (has a colon)
        let has_colon = sub_children.iter().any(|c| c.as_str().contains(':'));

        // Count expressions to differentiate
        let expr_children: Vec<&Pair<Rule>> = sub_children.iter()
            .filter(|c| c.as_rule() == Rule::expression)
            .collect();

        if has_colon || (expr_children.len() >= 1 && sub_children.iter().any(|c| c.as_str() == ":")) {
            // Reference modification: name(start:length)
            // → __refmod(name, start - 1, length)
            let start_expr = if !expr_children.is_empty() {
                walk_expression(expr_children[0].clone())?
            } else {
                Expression::int(1)
            };
            let length_expr = if expr_children.len() >= 2 {
                walk_expression(expr_children[1].clone())?
            } else {
                Expression::int(0) // No length = rest of string
            };

            // Adjust 1-indexed to 0-indexed: start - 1
            let adjusted_start = binary(BinOp::Sub, start_expr, Expression::int(1));

            expr = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__refmod")),
                args: vec![
                    Argument::positional(expr),
                    Argument::positional(adjusted_start),
                    Argument::positional(length_expr),
                ],
                optional: false,
            });
        } else if !expr_children.is_empty() {
            // Subscript: name(index) → name[index - 1]
            let index_expr = walk_expression(expr_children[0].clone())?;
            // COBOL is 1-indexed, subtract 1
            let adjusted_index = binary(BinOp::Sub, index_expr, Expression::int(1));
            expr = Expression::new(ExprKind::Index {
                object: Box::new(expr),
                index: Box::new(adjusted_index),
                null_safe: false,
            });

            // Handle multi-dimensional subscripts: name(i, j)
            for extra in expr_children.iter().skip(1) {
                let extra_idx = walk_expression((*extra).clone())?;
                let adjusted = binary(BinOp::Sub, extra_idx, Expression::int(1));
                expr = Expression::new(ExprKind::Index {
                    object: Box::new(expr),
                    index: Box::new(adjusted),
                    null_safe: false,
                });
            }
        }
    }

    // Handle qualification: field OF group → group.field
    if let Some(parent) = qualification {
        expr = Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&parent)),
            field: name,
            null_safe: false,
        });
    }

    Ok(expr)
}

// ════════════════════════════════════════════════════════════════════════════
// Literals
// ════════════════════════════════════════════════════════════════════════════

fn walk_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    let inner = pair.into_inner().next().ok_or("empty literal")?;

    match inner.as_rule() {
        Rule::figurative_constant => walk_figurative_constant(inner),
        Rule::boolean_literal => walk_boolean_literal(inner),
        Rule::number_literal => Ok(parse_number_literal(inner.as_str())),
        Rule::string_literal => walk_string_literal(&inner),
        _ => Err(format!("COBOL walker: unhandled literal rule {:?}", inner.as_rule())),
    }
}

fn walk_figurative_constant(pair: Pair<Rule>) -> Result<Expression, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Check for ALL "x" — repeat the character
    let has_all = children.iter().any(|c| c.as_rule() == Rule::kw_all);
    if has_all {
        // ALL "x" or ALL SPACES — the runtime will handle repetition
        for c in &children {
            if c.as_rule() == Rule::string_literal {
                return walk_string_literal(c);
            }
            if c.as_rule() == Rule::figurative_constant {
                return walk_figurative_constant(c.clone());
            }
        }
    }

    for child in &children {
        match child.as_rule() {
            Rule::kw_spaces | Rule::kw_space => {
                return Ok(Expression::string(" "));
            }
            Rule::kw_zeros | Rule::kw_zeroes | Rule::kw_zero => {
                return Ok(Expression::int(0));
            }
            Rule::kw_low_values | Rule::kw_low_value => {
                return Ok(Expression::string(""));
            }
            Rule::kw_high_values | Rule::kw_high_value => {
                return Ok(Expression::string("\u{FFFF}"));
            }
            Rule::kw_quotes | Rule::kw_quote => {
                return Ok(Expression::string("\""));
            }
            Rule::kw_nulls | Rule::kw_null => {
                return Ok(Expression::null());
            }
            _ => {}
        }
    }

    Ok(Expression::null())
}

fn walk_boolean_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    let inner = pair.into_inner().next().ok_or("empty boolean")?;
    match inner.as_rule() {
        Rule::kw_true => Ok(Expression::bool(true)),
        Rule::kw_false => Ok(Expression::bool(false)),
        _ => Ok(Expression::bool(false)),
    }
}

fn parse_number_literal(s: &str) -> Expression {
    let trimmed = s.trim();
    if trimmed.contains('.') {
        match trimmed.parse::<f64>() {
            Ok(f) => Expression::float(f),
            Err(_) => Expression::int(0),
        }
    } else {
        match trimmed.parse::<i64>() {
            Ok(n) => Expression::int(n),
            Err(_) => Expression::int(0),
        }
    }
}

fn walk_string_literal(pair: &Pair<Rule>) -> Result<Expression, String> {
    let raw = pair.as_str();
    // Strip surrounding quotes (either " or ')
    if raw.len() >= 2 {
        let inner = &raw[1..raw.len()-1];
        Ok(Expression::string(inner))
    } else {
        Ok(Expression::string(""))
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Helper functions
// ════════════════════════════════════════════════════════════════════════════

/// Create a binary expression.
fn binary(op: BinOp, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

/// Negate an expression (wrap in NOT).
fn negate_expr(expr: Expression) -> Expression {
    Expression::new(ExprKind::Unary {
        op: UnaryOp::Not,
        expr: Box::new(expr),
    })
}

/// Build a sum expression from a list of expressions.
fn build_sum_expr(exprs: &[Expression]) -> Expression {
    if exprs.is_empty() {
        return Expression::int(0);
    }
    let mut result = exprs[0].clone();
    for expr in &exprs[1..] {
        result = binary(BinOp::Add, result, expr.clone());
    }
    result
}
