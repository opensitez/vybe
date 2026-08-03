//! COBOL walker — pest `Pair<Rule>` → `vybe_compiler::ast::Module`.
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
//! - **1-indexed arrays**: COBOL arrays start at 1. The walker routes
//!   subscripts through the shared array normalization helper.
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

use super::{CobolParser, Rule};
use pest::Parser;
use pest::iterators::{Pair, Pairs};
use std::collections::{HashMap, HashSet};
use vybe_ast::*;

const COBOL_ARRAY_INDEXING: ArrayIndexSemantics = ArrayIndexSemantics::ONE_BASED;

#[derive(Clone)]
struct CobolRecordField {
    name: String,
    numeric: bool,
    decimal_places: usize }

#[derive(Clone)]
struct CobolFileBinding {
    path: Expression,
    file_number: i32,
    status_var: Option<String>,
    key_name: Option<String>,
    record_name: Option<String>,
    record_fields: Vec<CobolRecordField> }

struct CobolWalkerContext {
    file_bindings: HashMap<String, CobolFileBinding>,
    record_to_file: HashMap<String, String>,
    record_fields: HashMap<String, Vec<CobolRecordField>>,
    group_layouts: HashMap<String, Vec<Expression>>,
    condition_names: HashMap<String, Expression>,
    // Names declared in the SCREEN SECTION. DISPLAY of a screen item renders to
    // the terminal in real COBOL; here it is suppressed so it produces no stdout.
    screen_items: HashSet<String>,
    // Elementary working-storage fields whose PICTURE gives a fixed display
    // width, so DISPLAY of the field pads to that width.
    field_pics: HashMap<String, CobolPicFmt>,
    next_file_number: i32 }

/// Fixed-width DISPLAY format implied by an elementary field's PICTURE.
#[derive(Clone, Copy)]
enum CobolPicFmt {
    /// Unsigned integer PIC 9(n): zero-pad to `n` digits.
    Numeric(usize),
    /// Alphanumeric PIC X(n)/A(n): space-pad (right) to `n` characters.
    Alpha(usize) }

impl CobolWalkerContext {
    fn new() -> Self {
        Self {
            file_bindings: HashMap::new(),
            record_to_file: HashMap::new(),
            record_fields: HashMap::new(),
            group_layouts: HashMap::new(),
            condition_names: HashMap::new(),
            screen_items: HashSet::new(),
            field_pics: HashMap::new(),
            next_file_number: 1 }
    }

    fn register_screen_item(&mut self, name: &str) {
        self.screen_items.insert(cobol_name_key(name));
    }

    fn is_screen_item(&self, name: &str) -> bool {
        self.screen_items.contains(&cobol_name_key(name))
    }

    fn register_field_pic(&mut self, name: &str, pic: &str) {
        if let Some(fmt) = cobol_pic_display_fmt(pic) {
            self.field_pics.insert(cobol_name_key(name), fmt);
        }
    }

    fn field_pic(&self, name: &str) -> Option<CobolPicFmt> {
        self.field_pics.get(&cobol_name_key(name)).copied()
    }

    fn register_file_binding(
        &mut self,
        file_name: &str,
        path: Expression,
        status_var: Option<String>,
        key_name: Option<String>,
    ) {
        let key = cobol_name_key(file_name);
        let existing = self.file_bindings.get(&key);
        let record_name = self
            .file_bindings
            .get(&key)
            .and_then(|binding| binding.record_name.clone());
        let record_fields = self
            .file_bindings
            .get(&key)
            .map(|binding| binding.record_fields.clone())
            .unwrap_or_default();
        self.file_bindings.insert(
            key,
            CobolFileBinding {
                path,
                file_number: self.next_file_number,
                status_var,
                key_name: key_name
                    .or_else(|| existing.and_then(|binding| binding.key_name.clone())),
                record_name,
                record_fields },
        );
        self.next_file_number += 1;
    }

    fn bind_record_name(
        &mut self,
        file_name: &str,
        record_name: &str,
        record_fields: Vec<CobolRecordField>,
    ) {
        let file_key = cobol_name_key(file_name);
        let record_key = cobol_name_key(record_name);
        let record_fields = self.merge_record_fields(
            self.record_fields
                .get(&record_key)
                .map(Vec::as_slice)
                .unwrap_or(&[]),
            &record_fields,
        );
        self.record_to_file.insert(record_key, file_key.clone());
        self.record_fields
            .insert(cobol_name_key(record_name), record_fields.clone());
        if let Some(binding) = self.file_bindings.get_mut(&file_key) {
            binding.record_name = Some(record_name.to_string());
            binding.record_fields = record_fields;
        }
    }

    fn bind_record_fields_for_record(
        &mut self,
        record_name: &str,
        record_fields: Vec<CobolRecordField>,
    ) {
        let record_key = cobol_name_key(record_name);
        let record_fields = self.merge_record_fields(
            self.record_fields
                .get(&record_key)
                .map(Vec::as_slice)
                .unwrap_or(&[]),
            &record_fields,
        );
        self.record_fields
            .insert(record_key.clone(), record_fields.clone());
        let Some(file_key) = self.record_to_file.get(&record_key).cloned() else {
            return;
        };
        if let Some(binding) = self.file_bindings.get_mut(&file_key) {
            binding.record_fields = record_fields;
        }
    }

    fn merge_record_fields(
        &self,
        existing: &[CobolRecordField],
        incoming: &[CobolRecordField],
    ) -> Vec<CobolRecordField> {
        if existing.is_empty() {
            return incoming.to_vec();
        }

        incoming
            .iter()
            .map(|field| {
                let Some(previous) = existing
                    .iter()
                    .find(|existing_field| existing_field.name.eq_ignore_ascii_case(&field.name))
                else {
                    return field.clone();
                };

                CobolRecordField {
                    name: field.name.clone(),
                    numeric: field.numeric || previous.numeric,
                    decimal_places: field.decimal_places.max(previous.decimal_places) }
            })
            .collect()
    }

    fn bind_group_layout_for_name(&mut self, group_name: &str, layout: Vec<Expression>) {
        self.group_layouts
            .insert(cobol_name_key(group_name), layout);
    }

    fn record_fields_for_name(&self, record_name: &str) -> Vec<CobolRecordField> {
        self.record_fields
            .get(&cobol_name_key(record_name))
            .cloned()
            .unwrap_or_default()
    }

    fn group_layout_for_name(&self, group_name: &str) -> Vec<Expression> {
        self.group_layouts
            .get(&cobol_name_key(group_name))
            .cloned()
            .unwrap_or_default()
    }

    fn file_binding(&self, file_name: &str) -> Option<&CobolFileBinding> {
        self.file_bindings.get(&cobol_name_key(file_name))
    }

    fn file_binding_for_record(&self, record_name: &str) -> Option<&CobolFileBinding> {
        let file_key = self.record_to_file.get(&cobol_name_key(record_name))?;
        self.file_bindings.get(file_key)
    }

    fn register_condition_name(&mut self, name: &str, expr: Expression) {
        self.condition_names.insert(cobol_name_key(name), expr);
    }

    fn condition_expr(&self, name: &str) -> Option<Expression> {
        self.condition_names.get(&cobol_name_key(name)).cloned()
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Keyword filter
// ════════════════════════════════════════════════════════════════════════════

/// Returns true for `kw_*` token rules. Pest preserves atomic rule nodes as
/// siblings inside their parent rule's parse tree, so without this filter
/// the keyword tokens leak into walker positional indexing.
fn is_kw(r: Rule) -> bool {
    use Rule::*;
    matches!(
        r,
        kw_identification
            | kw_environment
            | kw_configuration
            | kw_data
            | kw_procedure
            | kw_division
            | kw_section
            | kw_program_id
            | kw_class_id
            | kw_interface_id
            | kw_method_id
            | kw_author
            | kw_date_written
            | kw_special_names
            | kw_repository
            | kw_input_output
            | kw_file_control
            | kw_decimal_point
            | kw_currency
            | kw_alphabet
            | kw_symbolic
            | kw_comma
            | kw_working_storage
            | kw_local_storage
            | kw_linkage
            | kw_file
            | kw_screen
            | kw_pic
            | kw_picture
            | kw_value
            | kw_occurs
            | kw_times
            | kw_depending
            | kw_redefines
            | kw_usage
            | kw_indexed
            | kw_filler
            | kw_blank
            | kw_justified
            | kw_just
            | kw_synchronized
            | kw_sync
            | kw_global
            | kw_external
            | kw_binary
            | kw_comp
            | kw_comp_3
            | kw_comp_5
            | kw_display_usage
            | kw_packed_decimal
            | kw_pointer
            | kw_index
            | kw_float_long
            | kw_float_short
            | kw_national
            | kw_boolean
            | kw_fd
            | kw_sd
            | kw_record
            | kw_block
            | kw_contains
            | kw_characters
            | kw_character
            | kw_label
            | kw_standard
            | kw_omitted
            | kw_spaces
            | kw_space
            | kw_zeros
            | kw_zeroes
            | kw_zero
            | kw_low_values
            | kw_low_value
            | kw_high_values
            | kw_high_value
            | kw_quotes
            | kw_quote
            | kw_nulls
            | kw_null
            | kw_all
            | kw_display
            | kw_accept
            | kw_move
            | kw_add
            | kw_subtract
            | kw_multiply
            | kw_divide
            | kw_cancel
            | kw_release
            | kw_return
            | kw_alter
            | kw_proceed
            | kw_compute
            | kw_if
            | kw_else
            | kw_then
            | kw_evaluate
            | kw_when
            | kw_other
            | kw_perform
            | kw_until
            | kw_varying
            | kw_thru
            | kw_through
            | kw_string
            | kw_unstring
            | kw_inspect
            | kw_tallying
            | kw_replacing
            | kw_converting
            | kw_leading
            | kw_trailing
            | kw_first
            | kw_initial
            | kw_common
            | kw_recursive
            | kw_call
            | kw_using
            | kw_returning
            | kw_initialize
            | kw_set
            | kw_go
            | kw_to
            | kw_stop
            | kw_run
            | kw_goback
            | kw_continue
            | kw_raise
            | kw_exception
            | kw_json
            | kw_xml
            | kw_generate
            | kw_parse
            | kw_processing
            | kw_encoding
            | kw_attributes
            | kw_xml_declaration
            | kw_namespace
            | kw_namespace_prefix
            | kw_suppress
            | kw_linage
            | kw_footing
            | kw_top
            | kw_bottom
            | kw_end_of_page
            | kw_open
            | kw_close
            | kw_read
            | kw_write
            | kw_rewrite
            | kw_delete
            | kw_start
            | kw_sort
            | kw_merge
            | kw_search
            | kw_copy
            | kw_invoke
            | kw_validate
            | kw_free
            | kw_allocate
            | kw_typedef
            | kw_exit
            | kw_not
            | kw_and
            | kw_or
            | kw_true
            | kw_false
            | kw_any
            | kw_with
            | kw_test
            | kw_before
            | kw_after
            | kw_async
            | kw_giving
            | kw_from
            | kw_by
            | kw_into
            | kw_on
            | kw_size
            | kw_error
            | kw_rounded
            | kw_truncation
            | kw_nearest_even
            | kw_nearest_toward_zero
            | kw_toward_greater
            | kw_toward_lesser
            | kw_away_from_zero
            | kw_prohibited
            | kw_remainder
            | kw_corresponding
            | kw_corr
            | kw_delimited
            | kw_delimiter
            | kw_count
            | kw_overflow
            | kw_input
            | kw_output
            | kw_extend
            | kw_i_o
            | kw_ascending
            | kw_descending
            | kw_key
            | kw_at
            | kw_end
            | kw_invalid
            | kw_next
            | kw_page
            | kw_advancing
            | kw_lines
            | kw_line
            | kw_column
            | kw_upon
            | kw_no
            | kw_auto
            | kw_required
            | kw_protected
            | kw_secure
            | kw_highlight
            | kw_reverse_video
            | kw_blink
            | kw_foreground_color
            | kw_background_color
            | kw_declaratives
            | kw_use
            | kw_debugging
            | kw_procedures
            | kw_numeric
            | kw_alphabetic
            | kw_alphabetic_lower
            | kw_alphabetic_upper
            | kw_positive
            | kw_negative
            | kw_equal
            | kw_greater
            | kw_less
            | kw_than
            | kw_not_less
            | kw_not_greater
            | kw_inherits
            | kw_implements
            | kw_factory
            | kw_object
            | kw_new
            | kw_self
            | kw_override
            | kw_property
            | kw_get
            | kw_is
            | kw_as
            | kw_wait
            | kw_for
            | kw_unit
            | kw_lock
            | kw_unlock
            | kw_yield
            | kw_suspend
            | kw_exec
            | kw_sql
            | kw_cics
            | kw_dli
            | kw_end_exec
            | kw_date
            | kw_time
            | kw_day
            | kw_day_of_week
            | kw_command_line
            | kw_console
            | kw_select
            | kw_assign
            | kw_organization
            | kw_sequential
            | kw_relative
            | kw_file_status
            | kw_access
            | kw_mode
            | kw_random
            | kw_dynamic
            | kw_alternate
            | kw_order
            | kw_duplicates
            | kw_up
            | kw_down
            | kw_source_computer
            | kw_object_computer
            | kw_reference
            | kw_content
            | kw_alphanumeric
            | kw_sign
            | kw_separate
            | kw_end_if
            | kw_end_evaluate
            | kw_end_perform
            | kw_end_call
            | kw_end_read
            | kw_end_write
            | kw_end_rewrite
            | kw_end_delete
            | kw_end_start
            | kw_end_string
            | kw_end_unstring
            | kw_end_search
            | kw_end_add
            | kw_end_subtract
            | kw_end_multiply
            | kw_end_divide
            | kw_end_return
            | kw_end_compute
            | kw_end_class
            | kw_end_method
            | kw_end_factory
            | kw_end_object
            | kw_end_interface
            | kw_end_validate
            | kw_end_json
            | kw_end_xml
            | kw_end_program
            | kw_name
            | kw_class
            | kw_paragraph
            | kw_program
            | kw_method
            | kw_cycle
            | kw_when_
            | kw_left
            | kw_right
            | kw_function
            | kw_length
            | kw_upper_case
            | kw_lower_case
            | kw_trim
            | kw_reverse
            | kw_current_date
            | kw_max
            | kw_min
            | kw_mod
            | kw_rem
            | kw_numval
            | kw_numval_c
            | kw_substitute
            | kw_sqrt
            | kw_sum
            | kw_integer
            | kw_abs
            | kw_ord
            | kw_char
            | kw_floor
            | kw_ceiling
            | kw_power
            | kw_log
            | kw_log10
            | kw_exp
            | kw_sin
            | kw_cos
            | kw_tan
            | kw_asin
            | kw_acos
            | kw_atan
            | kw_mean
            | kw_median
            | kw_variance
            | kw_concatenate
            | kw_when_compiled
            | kw_test_numval
            | kw_date_of_integer
            | kw_integer_of_date
            | kw_day_of_integer
            | kw_annuity
            | kw_present_value
            | kw_formatted_date
            | kw_formatted_time
            | kw_also
            | kw_in
            | kw_of
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
        end_col: end_col as u32 }
}

// ════════════════════════════════════════════════════════════════════════════
// Top-level entry point
// ════════════════════════════════════════════════════════════════════════════

pub fn parse(source: &str) -> Result<Module, String> {
    let mut pairs = CobolParser::parse(Rule::program, source)
        .map_err(|e| format!("COBOL parse error: {}", e))?;
    let program = pairs.next().ok_or("empty parse")?;
    let mut ctx = CobolWalkerContext::new();

    let mut module = Module {
        name: String::new(),
        language: Lang::Cobol,
        body: Vec::new(),
        imports: Vec::new() };

    for pair in program.into_inner() {
        match pair.as_rule() {
            Rule::identification_division => {
                walk_identification_division(pair, &mut module)?;
            }
            Rule::class_id_paragraph => {
                walk_class_id(pair, &mut module.body)?;
            }
            Rule::interface_id_paragraph => {
                walk_interface_id(pair, &mut module.body)?;
            }
            Rule::environment_division => {
                walk_environment_division(pair, &mut module, &mut ctx)?;
            }
            Rule::data_division => {
                walk_data_division(pair, &mut module.body, &mut ctx)?;
            }
            Rule::procedure_division => {
                walk_procedure_division(pair, &mut module.body, &ctx)?;
            }
            Rule::nested_program | Rule::EOI => {}
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
                if let Some(name_pair) = child
                    .into_inner()
                    .find(|p| matches!(p.as_rule(), Rule::ident_name | Rule::ident_or_keyword))
                {
                    module.name = name_pair.as_str().to_string();
                }
            }
            Rule::class_id_paragraph => {
                walk_class_id(child, &mut module.body)?;
            }
            Rule::interface_id_paragraph => {
                walk_interface_id(child, &mut module.body)?;
            }
            Rule::id_optional_paragraph => {}
            _ => {}
        }
    }
    Ok(())
}

// ════════════════════════════════════════════════════════════════════════════
// Environment Division
// ════════════════════════════════════════════════════════════════════════════

fn walk_environment_division(
    pair: Pair<Rule>,
    module: &mut Module,
    ctx: &mut CobolWalkerContext,
) -> Result<(), String> {
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::configuration_section => {
                for part in child.into_inner() {
                    if part.as_rule() != Rule::configuration_entry {
                        continue;
                    }
                    for entry in part.into_inner() {
                        if entry.as_rule() == Rule::repository_paragraph {
                            walk_repository_paragraph(entry, module)?;
                        }
                    }
                }
            }
            Rule::input_output_section => {
                for part in child.into_inner() {
                    if part.as_rule() != Rule::file_control_paragraph {
                        continue;
                    }
                    for entry in part.into_inner() {
                        if entry.as_rule() == Rule::select_entry {
                            walk_select_entry(entry, ctx)?;
                        }
                    }
                }
            }
            _ => {}
        }
    }
    Ok(())
}

fn walk_repository_paragraph(pair: Pair<Rule>, module: &mut Module) -> Result<(), String> {
    for entry in pair.into_inner() {
        if entry.as_rule() != Rule::repository_entry {
            continue;
        }
        for item in entry.into_inner() {
            match item.as_rule() {
                Rule::repository_function_entry => {
                    if let Some(import) = walk_repository_function_entry(item)? {
                        module.imports.push(import);
                    }
                }
                Rule::repository_class_entry | Rule::repository_interface_entry => {
                    if let Some(import) = walk_repository_type_entry(item)? {
                        module.imports.push(import);
                    }
                }
                _ => {}
            }
        }
    }
    Ok(())
}

fn walk_repository_function_entry(pair: Pair<Rule>) -> Result<Option<Import>, String> {
    let span = to_span(&pair);
    let mut local_name: Option<String> = None;
    let mut raw_specifier: Option<String> = None;
    let mut all_intrinsic = false;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::function_name if local_name.is_none() => {
                local_name = Some(child.as_str().to_string());
            }
            Rule::string_literal => {
                raw_specifier = Some(string_literal_value(&child));
            }
            Rule::kw_all => all_intrinsic = true,
            _ => {}
        }
    }

    if all_intrinsic {
        return Ok(None);
    }

    let Some(local_name) = local_name else {
        return Ok(None);
    };
    let Some(raw_specifier) = raw_specifier else {
        return Ok(None);
    };

    let (path, imported_name) = split_repository_member_spec(&raw_specifier, &local_name);
    let alias = if imported_name.eq_ignore_ascii_case(&local_name) {
        None
    } else {
        Some(local_name)
    };

    Ok(Some(Import {
        kind: ImportKind::Named {
            path,
            names: vec![ImportName {
                name: imported_name,
                alias }],
            level: 0 },
        span }))
}

fn walk_repository_type_entry(pair: Pair<Rule>) -> Result<Option<Import>, String> {
    let span = to_span(&pair);
    let mut alias: Option<String> = None;
    let mut target: Option<String> = None;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::ident_name if alias.is_none() => alias = Some(child.as_str().to_string()),
            Rule::string_literal => target = Some(string_literal_value(&child)),
            _ => {}
        }
    }

    let (Some(alias), Some(target)) = (alias, target) else {
        return Ok(None);
    };

    Ok(Some(Import {
        kind: ImportKind::Simple {
            path: target,
            alias: Some(alias) },
        span }))
}

fn split_repository_member_spec(raw_specifier: &str, fallback_name: &str) -> (String, String) {
    if let Some((path, name)) = raw_specifier.rsplit_once('#') {
        if !path.is_empty() && !name.is_empty() {
            return (path.to_string(), name.to_string());
        }
    }
    if let Some((path, name)) = raw_specifier.rsplit_once("::") {
        if !path.is_empty() && !name.is_empty() {
            return (path.to_string(), name.to_string());
        }
    }
    if let Some((path, name)) = split_host_repository_spec(raw_specifier) {
        return (path, name);
    }
    (raw_specifier.to_string(), fallback_name.to_string())
}

fn split_host_repository_spec(raw_specifier: &str) -> Option<(String, String)> {
    let is_host_like = raw_specifier.starts_with("wasi:")
        || raw_specifier.starts_with("wasm:")
        || raw_specifier.starts_with("vybe:")
        || raw_specifier.starts_with("node:");
    if !is_host_like {
        return None;
    }

    let split_at = raw_specifier.rfind(':')?;
    let path = &raw_specifier[..split_at];
    let name = &raw_specifier[split_at + 1..];
    if path.is_empty() || name.is_empty() || name.contains('/') {
        return None;
    }
    Some((path.to_string(), name.to_string()))
}

fn walk_select_entry(pair: Pair<Rule>, ctx: &mut CobolWalkerContext) -> Result<(), String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    let mut file_name = String::new();
    let mut assign_target: Option<Expression> = None;
    let mut status_var: Option<String> = None;
    let mut key_name: Option<String> = None;

    for child in children {
        match child.as_rule() {
            Rule::ident_name if file_name.is_empty() => {
                file_name = child.as_str().to_string();
            }
            Rule::string_literal if assign_target.is_none() => {
                assign_target = Some(walk_string_literal(&child)?);
            }
            Rule::ident_name if assign_target.is_none() => {
                assign_target = Some(Expression::ident(child.as_str()));
            }
            Rule::select_clause => {
                let clause_children: Vec<Pair<Rule>> = child.into_inner().collect();
                if clause_children
                    .iter()
                    .any(|part| part.as_rule() == Rule::kw_file_status)
                {
                    status_var = clause_children
                        .iter()
                        .find(|part| part.as_rule() == Rule::ident_name)
                        .map(|part| part.as_str().to_string());
                }
                if !clause_children
                    .iter()
                    .any(|part| part.as_rule() == Rule::kw_alternate)
                    && clause_children
                        .iter()
                        .any(|part| matches!(part.as_rule(), Rule::kw_relative | Rule::kw_record))
                    && clause_children
                        .iter()
                        .any(|part| part.as_rule() == Rule::kw_key)
                {
                    key_name = clause_children
                        .iter()
                        .rev()
                        .find(|part| part.as_rule() == Rule::ident_name)
                        .map(|part| part.as_str().to_string());
                }
            }
            _ => {}
        }
    }

    if !file_name.is_empty() {
        ctx.register_file_binding(
            &file_name,
            assign_target.unwrap_or_else(|| Expression::string(&file_name.to_ascii_lowercase())),
            status_var,
            key_name,
        );
    }
    Ok(())
}

fn extract_cobol_data_item_name(pair: &Pair<Rule>) -> Option<String> {
    if pair.as_rule() != Rule::data_item {
        return None;
    }
    pair.clone()
        .into_inner()
        .find_map(|child| match child.as_rule() {
            Rule::regular_data_item => child.into_inner().find_map(|part| match part.as_rule() {
                Rule::ident_name | Rule::ident_or_keyword => Some(part.as_str().to_string()),
                _ => None }),
            _ => None })
}

fn collect_cobol_record_fields(pair: Pair<Rule>, fields: &mut Vec<CobolRecordField>) {
    match pair.as_rule() {
        Rule::data_item => {
            for child in pair.into_inner() {
                collect_cobol_record_fields(child, fields);
            }
        }
        Rule::regular_data_item => {
            let level = cobol_data_item_level(&pair);
            let mut name = String::new();
            let mut nested_items = Vec::new();
            let mut pic_str: Option<String> = None;
            let mut usage_str: Option<String> = None;

            for child in pair.into_inner() {
                match child.as_rule() {
                    Rule::ident_name | Rule::ident_or_keyword => {
                        if name.is_empty() {
                            name = child.as_str().to_string();
                        }
                    }
                    Rule::kw_filler => {
                        name = "FILLER".to_string();
                    }
                    Rule::data_clause
                    | Rule::pic_clause
                    | Rule::usage_clause
                    | Rule::value_clause
                    | Rule::occurs_clause => {
                        let mut ignored_init = None;
                        let mut ignored_occurs = None;
                        let _ = walk_data_clause(
                            child,
                            &mut pic_str,
                            &mut usage_str,
                            &mut ignored_init,
                            &mut ignored_occurs,
                        );
                    }
                    Rule::data_item => nested_items.push(child),
                    _ => {}
                }
            }

            let (true_children, leaked_siblings): (Vec<_>, Vec<_>) = nested_items
                .into_iter()
                .partition(|item| cobol_data_item_level(item) > level);

            if !true_children.is_empty() {
                for nested in true_children {
                    collect_cobol_record_fields(nested, fields);
                }
            } else if !name.is_empty() && name != "FILLER" {
                fields.push(CobolRecordField {
                    name,
                    numeric: cobol_type_hint(pic_str.as_deref(), usage_str.as_deref())
                        .as_deref()
                        .is_some_and(|hint| hint != "String" && hint != "Boolean"),
                    decimal_places: pic_fractional_digits(pic_str.as_deref().unwrap_or_default()) });
            }

            for sibling in leaked_siblings {
                collect_cobol_record_fields(sibling, fields);
            }
        }
        _ => {}
    }
}

fn cobol_file_status_assign(status_var: Option<&str>, value: &str) -> Option<Statement> {
    status_var.map(|name| {
        Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident(name)],
            value: Expression::string(value), by_ref: false })
    })
}

fn cobol_record_key_index(
    binding: &CobolFileBinding,
    record_fields: &[CobolRecordField],
    key_name_override: Option<&str>,
) -> usize {
    key_name_override
        .or(binding.key_name.as_deref())
        .and_then(|key_name| {
            record_fields
                .iter()
                .position(|field| field.name.eq_ignore_ascii_case(key_name))
        })
        .unwrap_or(0)
}

fn parse_cobol_start_relation(source: &str) -> FileKeyRelation {
    let lower = source.to_ascii_lowercase();
    if lower.contains("not less") || lower.contains(">=") {
        FileKeyRelation::GreaterOrEqual
    } else if lower.contains("not greater") || lower.contains("<=") {
        FileKeyRelation::LessOrEqual
    } else if lower.contains(" > ") {
        FileKeyRelation::Greater
    } else if lower.contains(" < ") {
        FileKeyRelation::Less
    } else {
        FileKeyRelation::Equal
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Data Division
// ════════════════════════════════════════════════════════════════════════════

fn walk_data_division(
    pair: Pair<Rule>,
    body: &mut Vec<Statement>,
    ctx: &mut CobolWalkerContext,
) -> Result<(), String> {
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::file_section => {
                walk_file_section(child, body, ctx)?;
            }
            Rule::working_storage_section | Rule::local_storage_section | Rule::linkage_section => {
                walk_storage_section(child, body, ctx)?;
            }
            Rule::screen_section => {
                walk_screen_section(child, body, ctx)?;
            }
            _ => {}
        }
    }
    Ok(())
}

fn walk_file_section(
    pair: Pair<Rule>,
    body: &mut Vec<Statement>,
    ctx: &mut CobolWalkerContext,
) -> Result<(), String> {
    for child in pair.into_inner() {
        if child.as_rule() == Rule::file_description {
            walk_file_description(child, body, ctx)?;
        }
    }
    Ok(())
}

fn walk_file_description(
    pair: Pair<Rule>,
    body: &mut Vec<Statement>,
    ctx: &mut CobolWalkerContext,
) -> Result<(), String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    let file_name = children
        .iter()
        .find(|child| child.as_rule() == Rule::ident_or_keyword)
        .map(|child| child.as_str().to_string())
        .unwrap_or_default();
    let data_items: Vec<Pair<Rule>> = children
        .iter()
        .filter(|child| child.as_rule() == Rule::data_item)
        .cloned()
        .collect();
    if let Some(record_item) = data_items.first().cloned() {
        let record_name = extract_cobol_data_item_name(&record_item).unwrap_or_default();
        let mut record_fields = Vec::new();
        for item in data_items.iter().skip(1) {
            collect_cobol_record_fields(item.clone(), &mut record_fields);
        }
        if record_fields.is_empty() {
            collect_cobol_record_fields(record_item, &mut record_fields);
        }
        record_fields.dedup_by(|a, b| a.name.eq_ignore_ascii_case(&b.name));
        if !file_name.is_empty() {
            ctx.bind_record_name(&file_name, &record_name, record_fields);
        }
    }
    for child in children {
        if child.as_rule() == Rule::data_item {
            walk_data_item(child, body, ctx)?;
        }
    }
    Ok(())
}

fn walk_storage_section(
    pair: Pair<Rule>,
    body: &mut Vec<Statement>,
    ctx: &mut CobolWalkerContext,
) -> Result<(), String> {
    for child in pair.into_inner() {
        if child.as_rule() == Rule::data_item {
            walk_data_item(child, body, ctx)?;
        }
    }
    Ok(())
}

fn walk_screen_section(
    pair: Pair<Rule>,
    body: &mut Vec<Statement>,
    ctx: &mut CobolWalkerContext,
) -> Result<(), String> {
    for child in pair.into_inner() {
        if child.as_rule() == Rule::screen_item {
            walk_screen_item(child, body, ctx)?;
        }
    }
    Ok(())
}

fn walk_screen_item(
    pair: Pair<Rule>,
    body: &mut Vec<Statement>,
    ctx: &mut CobolWalkerContext,
) -> Result<(), String> {
    let span = to_span(&pair);
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::level_88_item => {
                let stmt = walk_level_88(child)?;
                body.push(Statement::with_span(stmt, span));
            }
            Rule::screen_data_item => {
                walk_screen_data_item(child, body, ctx)?;
            }
            Rule::screen_item => {
                walk_screen_item(child, body, ctx)?;
            }
            _ => {}
        }
    }
    Ok(())
}

fn walk_screen_data_item(
    pair: Pair<Rule>,
    body: &mut Vec<Statement>,
    ctx: &mut CobolWalkerContext,
) -> Result<(), String> {
    let span = to_span(&pair);
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut name = String::new();
    let mut pic_str: Option<String> = None;
    let mut usage_str: Option<String> = None;
    let mut init_value: Option<Expression> = None;
    let mut nested_items: Vec<Pair<Rule>> = Vec::new();

    for child in children {
        match child.as_rule() {
            Rule::screen_item_name => {
                if name.is_empty() {
                    name = child.as_str().to_string();
                }
            }
            Rule::pic_clause => {
                for part in child.into_inner() {
                    if part.as_rule() == Rule::pic_string {
                        pic_str = Some(part.as_str().to_string());
                    }
                }
            }
            Rule::usage_clause => {
                usage_str = extract_usage_clause_name(&child);
            }
            Rule::value_clause => {
                for part in inner_nokw(child) {
                    if part.as_rule() == Rule::literal {
                        init_value = Some(walk_literal(part)?);
                    }
                }
            }
            Rule::screen_item => nested_items.push(child),
            Rule::screen_source_clause
            | Rule::blank_screen_clause
            | Rule::line_clause
            | Rule::column_clause
            | Rule::highlight_clause
            | Rule::reverse_video_clause
            | Rule::blink_clause
            | Rule::auto_clause
            | Rule::required_clause
            | Rule::protected_clause
            | Rule::secure_clause
            | Rule::foreground_color_clause
            | Rule::background_color_clause
            | Rule::level_number
            | Rule::period => {}
            _ => {
                if is_kw(child.as_rule()) {
                    continue;
                }
            }
        }
    }

    for nested in nested_items {
        walk_screen_item(nested, body, ctx)?;
    }

    if name.is_empty() {
        return Ok(());
    }

    ctx.register_screen_item(&name);

    body.push(Statement::with_span(
        StmtKind::VarDecl {
            kind: VarDeclKind::Dim,
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(name),
                type_hint: cobol_type_hint(pic_str.as_deref(), usage_str.as_deref()),
                init: init_value.or_else(|| {
                    Some(default_value_for_cobol_type(
                        pic_str.as_deref(),
                        usage_str.as_deref(),
                    ))
                }),
                array_bounds: None,
                with_events: false }] },
        span,
    ));
    Ok(())
}

// ── Data Items ─────────────────────────────────────────────────────────────

fn walk_data_item(
    pair: Pair<Rule>,
    body: &mut Vec<Statement>,
    ctx: &mut CobolWalkerContext,
) -> Result<(), String> {
    let span = to_span(&pair);
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::level_88_item => {
                // 88-level condition → constant declaration
                let stmt = walk_level_88(child)?;
                body.push(Statement::with_span(stmt, span));
            }
            Rule::regular_data_item => {
                walk_regular_data_item(child, body, ctx)?;
            }
            Rule::data_item => {
                // Nested data items (children of a group)
                walk_data_item(child, body, ctx)?;
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
            with_events: false }] })
}

fn register_level_88_condition(
    pair: Pair<Rule>,
    parent_name: &str,
    ctx: &mut CobolWalkerContext,
) -> Result<(), String> {
    if parent_name.is_empty() {
        return Ok(());
    }

    let parts = inner_nokw(pair);
    let mut name = String::new();
    let mut values = Vec::new();

    for part in parts {
        match part.as_rule() {
            Rule::ident_name => {
                name = part.as_str().to_string();
            }
            Rule::literal => values.push(walk_literal(part)?),
            _ => {}
        }
    }

    if name.is_empty() || values.is_empty() {
        return Ok(());
    }

    let mut comparisons = values
        .into_iter()
        .map(|value| binary(BinOp::Eq, Expression::ident(parent_name), value));
    let first = comparisons.next().unwrap();
    let expr = comparisons.fold(first, |acc, cmp| binary(BinOp::Or, acc, cmp));
    ctx.register_condition_name(&name, expr);
    Ok(())
}

fn walk_regular_data_item(
    pair: Pair<Rule>,
    body: &mut Vec<Statement>,
    ctx: &mut CobolWalkerContext,
) -> Result<(), String> {
    let span = to_span(&pair);
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut level: u32 = 0;
    let mut name = String::new();
    let mut pic_str: Option<String> = None;
    let mut usage_str: Option<String> = None;
    let mut init_value: Option<Expression> = None;
    let mut occurs_count: Option<Expression> = None;
    let mut is_filler = false;
    let mut nested_items: Vec<Pair<Rule>> = Vec::new();

    for child in children {
        match child.as_rule() {
            Rule::level_number => {
                level = child.as_str().trim().parse::<u32>().unwrap_or(0);
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
            Rule::data_clause
            | Rule::pic_clause
            | Rule::usage_clause
            | Rule::value_clause
            | Rule::occurs_clause => {
                walk_data_clause(
                    child,
                    &mut pic_str,
                    &mut usage_str,
                    &mut init_value,
                    &mut occurs_count,
                )?;
            }
            Rule::level_88_item => {
                register_level_88_condition(child, &name, ctx)?;
            }
            Rule::data_item => {
                nested_items.push(child);
            }
            Rule::period => {}
            _ => {
                if is_kw(child.as_rule()) {
                    continue;
                }
            }
        }
    }

    // Skip FILLER items (they are padding)
    if is_filler {
        // Still process nested items
        for nested in nested_items {
            walk_data_item(nested, body, ctx)?;
        }
        return Ok(());
    }

    let (group_children, sibling_items): (Vec<_>, Vec<_>) = nested_items
        .into_iter()
        .partition(|item| cobol_data_item_level(item) > level);

    // Determine type hint from PIC
    let type_hint = cobol_type_hint(pic_str.as_deref(), usage_str.as_deref());

    // A plain scalar elementary field (not a group, not an OCCURS table) whose
    // PICTURE gives a fixed display width — record it so DISPLAY of the field pads.
    let is_scalar_elementary = group_children.is_empty() && occurs_count.is_none();

    // Determine initial value
    let init = if !group_children.is_empty() {
        // Group item → Object initialiser with child fields
        let mut props = Vec::new();
        let mut child_stmts = Vec::new();
        let mut layout_parts = Vec::new();
        for nested in group_children.iter().chain(sibling_items.iter()) {
            // Walk nested item and collect as object property
            collect_group_children(nested.clone(), &mut props, &mut child_stmts)?;
            collect_group_layout_parts(nested.clone(), &mut layout_parts)?;
        }
        let mut field_names = Vec::new();
        collect_object_record_fields(&props, &mut field_names);
        if !field_names.is_empty() {
            ctx.bind_record_fields_for_record(&name, field_names);
        }
        if !layout_parts.is_empty() {
            ctx.bind_group_layout_for_name(&name, layout_parts);
        }
        // Group children are also directly addressable in COBOL, so emit
        // standalone bindings for them in addition to the parent object.
        for nested in group_children {
            walk_data_item(nested, body, ctx)?;
        }
        if props.is_empty() {
            init_value
        } else {
            Some(Expression::new(ExprKind::Object(props)))
        }
    } else if let Some(count_expr) = occurs_count {
        // OCCURS → array initialiser
        let element_init = init_value.clone().unwrap_or_else(|| {
            default_value_for_cobol_type(pic_str.as_deref(), usage_str.as_deref())
        });
        Some(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("Array")),
            args: vec![
                Argument::positional(count_expr),
                Argument::positional(element_init),
            ],
            optional: false }))
    } else {
        init_value.or_else(|| {
            Some(default_value_for_cobol_type(
                pic_str.as_deref(),
                usage_str.as_deref(),
            ))
        })
    };

    if name.is_empty() {
        return Ok(());
    }

    if is_scalar_elementary {
        if let Some(pic) = pic_str.as_deref() {
            ctx.register_field_pic(&name, pic);
        }
    }

    let stmt = StmtKind::VarDecl {
        kind: VarDeclKind::Dim,
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name),
            type_hint,
            init,
            array_bounds: None,
            with_events: false }] };
    body.push(Statement::with_span(stmt, span));
    for sibling in sibling_items {
        walk_data_item(sibling, body, ctx)?;
    }
    Ok(())
}

fn walk_data_clause(
    pair: Pair<Rule>,
    pic_str: &mut Option<String>,
    usage_str: &mut Option<String>,
    init_value: &mut Option<Expression>,
    occurs_count: &mut Option<Expression>,
) -> Result<(), String> {
    match pair.as_rule() {
        Rule::pic_clause => {
            for p in pair.into_inner() {
                if p.as_rule() == Rule::pic_string {
                    *pic_str = Some(p.as_str().to_string());
                }
            }
            return Ok(());
        }
        Rule::usage_clause => {
            *usage_str = extract_usage_clause_name(&pair);
            return Ok(());
        }
        Rule::value_clause => {
            let parts = inner_nokw(pair);
            for p in parts {
                if matches!(
                    p.as_rule(),
                    Rule::literal
                        | Rule::string_literal
                        | Rule::number_literal
                        | Rule::figurative_constant
                        | Rule::boolean_literal
                ) {
                    *init_value = Some(match p.as_rule() {
                        Rule::literal => walk_literal(p)?,
                        Rule::string_literal => walk_string_literal(&p)?,
                        Rule::number_literal => parse_number_literal(p.as_str()),
                        Rule::figurative_constant => walk_figurative_constant(p)?,
                        Rule::boolean_literal => walk_boolean_literal(p)?,
                        _ => unreachable!() });
                } else if matches!(p.as_rule(), Rule::ident_name | Rule::ident_or_keyword) {
                    let name = p.as_str();
                    if name.eq_ignore_ascii_case("true") {
                        *init_value = Some(Expression::bool(true));
                    } else if name.eq_ignore_ascii_case("false") {
                        *init_value = Some(Expression::bool(false));
                    } else {
                        *init_value = Some(Expression::ident(name));
                    }
                }
            }
            return Ok(());
        }
        Rule::occurs_clause => {
            let parts = inner_nokw(pair);
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
            return Ok(());
        }
        _ => {}
    }

    if pair
        .as_str()
        .trim_start()
        .to_ascii_uppercase()
        .starts_with("USAGE")
    {
        *usage_str = extract_usage_clause_name(&pair);
        return Ok(());
    }

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::pic_clause => {
                for p in child.into_inner() {
                    if p.as_rule() == Rule::pic_string {
                        *pic_str = Some(p.as_str().to_string());
                    }
                }
            }
            Rule::usage_clause => {
                *usage_str = extract_usage_clause_name(&child);
            }
            Rule::value_clause => {
                let parts = inner_nokw(child);
                for p in parts {
                    if matches!(
                        p.as_rule(),
                        Rule::literal
                            | Rule::string_literal
                            | Rule::number_literal
                            | Rule::figurative_constant
                            | Rule::boolean_literal
                    ) {
                        *init_value = Some(match p.as_rule() {
                            Rule::literal => walk_literal(p)?,
                            Rule::string_literal => walk_string_literal(&p)?,
                            Rule::number_literal => parse_number_literal(p.as_str()),
                            Rule::figurative_constant => walk_figurative_constant(p)?,
                            Rule::boolean_literal => walk_boolean_literal(p)?,
                            _ => unreachable!() });
                    } else if matches!(p.as_rule(), Rule::ident_name | Rule::ident_or_keyword) {
                        let name = p.as_str();
                        if name.eq_ignore_ascii_case("true") {
                            *init_value = Some(Expression::bool(true));
                        } else if name.eq_ignore_ascii_case("false") {
                            *init_value = Some(Expression::bool(false));
                        } else {
                            *init_value = Some(Expression::ident(name));
                        }
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
            Rule::redefines_clause
            | Rule::blank_clause
            | Rule::justified_clause
            | Rule::sign_clause
            | Rule::sync_clause
            | Rule::global_clause
            | Rule::external_clause
            | Rule::national_clause => {
                // These affect storage layout but not the AST structure
            }
            _ => {}
        }
    }
    Ok(())
}

fn extract_usage_clause_name(pair: &Pair<Rule>) -> Option<String> {
    pair.as_str()
        .split_whitespace()
        .last()
        .map(|token| token.trim().to_string())
        .filter(|token| !token.is_empty())
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
                let mut field_level: u32 = 0;
                let mut field_name = String::new();
                let mut field_pic: Option<String> = None;
                let mut field_usage: Option<String> = None;
                let mut field_init: Option<Expression> = None;
                let mut field_occurs: Option<Expression> = None;
                let mut sub_items: Vec<Pair<Rule>> = Vec::new();

                for c in children {
                    match c.as_rule() {
                        Rule::level_number => {
                            field_level = c.as_str().trim().parse::<u32>().unwrap_or(0);
                        }
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
                        Rule::data_clause
                        | Rule::pic_clause
                        | Rule::usage_clause
                        | Rule::value_clause
                        | Rule::occurs_clause => {
                            walk_data_clause(
                                c,
                                &mut field_pic,
                                &mut field_usage,
                                &mut field_init,
                                &mut field_occurs,
                            )?;
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

                let (true_children, leaked_siblings): (Vec<_>, Vec<_>) = sub_items
                    .into_iter()
                    .partition(|item| cobol_data_item_level(item) > field_level);

                let value = if !true_children.is_empty() {
                    let mut sub_props = Vec::new();
                    for si in true_children {
                        collect_group_children(si, &mut sub_props, extra_stmts)?;
                    }
                    Expression::new(ExprKind::Object(sub_props))
                } else if let Some(count_expr) = field_occurs {
                    let element_init = field_init.unwrap_or_else(|| {
                        default_value_for_cobol_type(field_pic.as_deref(), field_usage.as_deref())
                    });
                    eprintln!(
                        "[D1 WALKER] nested OCCURS path emitting Array(count, init) for field"
                    );
                    Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("Array")),
                        args: vec![
                            Argument::positional(count_expr),
                            Argument::positional(element_init),
                        ],
                        optional: false })
                } else {
                    field_init.unwrap_or_else(|| {
                        default_value_for_cobol_type(field_pic.as_deref(), field_usage.as_deref())
                    })
                };

                props.push(ObjectProperty::KeyValue {
                    key: Expression::string(&field_name),
                    value });

                for sibling in leaked_siblings {
                    collect_group_children(sibling, props, extra_stmts)?;
                }
            }
            Rule::data_item => {
                collect_group_children(child, props, extra_stmts)?;
            }
            _ => {}
        }
    }
    Ok(())
}

fn collect_group_layout_parts(pair: Pair<Rule>, parts: &mut Vec<Expression>) -> Result<(), String> {
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::regular_data_item => {
                let children: Vec<Pair<Rule>> = child.into_inner().collect();
                let mut field_level: u32 = 0;
                let mut field_name = String::new();
                let mut field_pic: Option<String> = None;
                let mut field_usage: Option<String> = None;
                let mut field_init: Option<Expression> = None;
                let mut sub_items: Vec<Pair<Rule>> = Vec::new();

                for c in children {
                    match c.as_rule() {
                        Rule::level_number => {
                            field_level = c.as_str().trim().parse::<u32>().unwrap_or(0);
                        }
                        Rule::ident_name | Rule::ident_or_keyword => {
                            if field_name.is_empty() {
                                field_name = c.as_str().to_string();
                            }
                        }
                        Rule::kw_filler => {
                            field_name = "FILLER".to_string();
                        }
                        Rule::data_clause
                        | Rule::pic_clause
                        | Rule::usage_clause
                        | Rule::value_clause
                        | Rule::occurs_clause => {
                            let mut ignored_occurs = None;
                            walk_data_clause(
                                c,
                                &mut field_pic,
                                &mut field_usage,
                                &mut field_init,
                                &mut ignored_occurs,
                            )?;
                        }
                        Rule::data_item => {
                            sub_items.push(c);
                        }
                        Rule::level_88_item | Rule::period => {}
                        _ => {}
                    }
                }

                let (true_children, leaked_siblings): (Vec<_>, Vec<_>) = sub_items
                    .into_iter()
                    .partition(|item| cobol_data_item_level(item) > field_level);

                if !true_children.is_empty() {
                    for nested in true_children {
                        collect_group_layout_parts(nested, parts)?;
                    }
                } else if field_name == "FILLER" || field_name.is_empty() {
                    parts.push(field_init.unwrap_or_else(|| {
                        default_value_for_cobol_type(field_pic.as_deref(), field_usage.as_deref())
                    }));
                } else {
                    parts.push(Expression::ident(&field_name));
                }

                for sibling in leaked_siblings {
                    collect_group_layout_parts(sibling, parts)?;
                }
            }
            Rule::data_item => {
                collect_group_layout_parts(child, parts)?;
            }
            _ => {}
        }
    }
    Ok(())
}

fn collect_object_record_fields(props: &[ObjectProperty], fields: &mut Vec<CobolRecordField>) {
    for prop in props {
        if let ObjectProperty::KeyValue { key, value } = prop {
            let ExprKind::Lit(Literal::Str(name)) = &key.kind else {
                continue;
            };

            if let ExprKind::Object(children) = &value.kind {
                collect_object_record_fields(children, fields);
                continue;
            }

            fields.push(CobolRecordField {
                name: name.clone(),
                numeric: !matches!(value.kind, ExprKind::Lit(Literal::Str(_))),
                decimal_places: 0 });
        }
    }
}

fn cobol_type_hint(pic: Option<&str>, usage: Option<&str>) -> Option<String> {
    if let Some(usage) = usage {
        let usage = usage.trim().to_ascii_uppercase();
        let hint = match usage.as_str() {
            "BOOLEAN" => Some("Boolean"),
            "FLOAT-SHORT" => Some("Single"),
            "FLOAT-LONG" => Some("Double"),
            "COMP-3" | "PACKED-DECIMAL" => Some("Decimal"),
            "POINTER" | "INDEX" => Some("Long"),
            "NATIONAL" => Some("String"),
            "BINARY" | "COMP" | "COMP-5" => {
                Some(if pic_integer_digits(pic.unwrap_or_default()) > 9 {
                    "Long"
                } else {
                    "Integer"
                })
            }
            _ => None };
        if let Some(hint) = hint {
            return Some(hint.to_string());
        }
    }

    let pic = pic?;
    let upper = pic.trim().to_ascii_uppercase();
    if upper.is_empty() {
        return None;
    }
    if upper.starts_with('X') || upper.starts_with('A') || upper.starts_with('N') {
        return Some("String".to_string());
    }
    if upper.starts_with('B') && !upper.contains('9') {
        return Some("Boolean".to_string());
    }
    if upper.contains('V') || upper.contains('.') || upper.contains('P') {
        return Some("Decimal".to_string());
    }
    if pic_integer_digits(&upper) > 9 {
        return Some("Long".to_string());
    }
    Some("Integer".to_string())
}

fn default_value_for_cobol_type(pic: Option<&str>, usage: Option<&str>) -> Expression {
    match cobol_type_hint(pic, usage).as_deref() {
        Some("String") => Expression::string(" "),
        Some("Boolean") => Expression::bool(false),
        Some("Integer") | Some("Long") | Some("Single") | Some("Double") | Some("Decimal") => {
            Expression::int(0)
        }
        Some(_) | None => Expression::null() }
}

/// Fixed-width DISPLAY format for an elementary PICTURE, if it is a plain
/// unsigned integer (`9(n)`) or alphanumeric (`X(n)`/`A(n)`). Signed, decimal
/// (`V`/`.`), and numeric-edited pictures (`Z * + - , $ B / CR DB`) return None —
/// those need the full editing engine and are handled separately.
/// Wrap a field value in the pad call implied by its PICTURE for DISPLAY:
/// numeric 9(n) → `padStart(toFixed(v,0), n, "0")`; alphanumeric X(n) →
/// `padEnd(v, n, " ")`. Uses the shared ECMA string helpers.
fn cobol_pic_format_expr(value: Expression, fmt: CobolPicFmt) -> Expression {
    let call = |fname: &str, args: Vec<Expression>| -> Expression {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(fname)),
            args: args.into_iter().map(Argument::positional).collect(),
            optional: false })
    };
    match fmt {
        CobolPicFmt::Numeric(digits) => {
            let int_str = call("__to_fixed2", vec![value, Expression::int(0)]);
            call(
                "__pad_start",
                vec![
                    int_str,
                    Expression::int(digits as i64),
                    Expression::string("0"),
                ],
            )
        }
        CobolPicFmt::Alpha(width) => call(
            "__pad_end",
            vec![
                value,
                Expression::int(width as i64),
                Expression::string(" "),
            ],
        ) }
}

fn cobol_pic_display_fmt(pic: &str) -> Option<CobolPicFmt> {
    let upper = pic.trim().to_ascii_uppercase();
    if upper.is_empty() {
        return None;
    }
    // Alphanumeric first: any X marker → space-padded text field.
    if upper.contains('X') {
        let width = count_pic_markers_before_decimal(&upper, &['X', 'A', 'N']);
        return (width > 0).then_some(CobolPicFmt::Alpha(width));
    }
    // Reject signed / decimal / edited pictures.
    if upper.contains(['S', 'V', '.', 'Z', '*', '+', '-', ',', '$', 'B', '/', 'P']) {
        return None;
    }
    // Only 9 digit positions (with optional `(n)` repeat) qualify as plain numeric.
    if upper
        .chars()
        .all(|c| matches!(c, '9' | '(' | ')') || c.is_ascii_digit())
    {
        let digits = count_pic_markers_before_decimal(&upper, &['9']);
        return (digits > 0).then_some(CobolPicFmt::Numeric(digits));
    }
    None
}

fn pic_integer_digits(pic: &str) -> usize {
    count_pic_markers_before_decimal(pic, &['9', 'Z', '*'])
}

fn pic_fractional_digits(pic: &str) -> usize {
    count_pic_markers_after_decimal(pic, &['9', 'Z', '*'])
}

fn count_pic_markers_before_decimal(pic: &str, markers: &[char]) -> usize {
    let upper = pic.trim().to_ascii_uppercase();
    let chars: Vec<char> = upper.chars().collect();
    let mut index = 0;
    let mut count = 0;

    while index < chars.len() {
        let ch = chars[index];
        if ch == 'V' || ch == '.' {
            break;
        }
        if markers.contains(&ch) {
            let mut repeat = 1usize;
            if index + 1 < chars.len() && chars[index + 1] == '(' {
                let mut cursor = index + 2;
                let mut digits = String::new();
                while cursor < chars.len() && chars[cursor] != ')' {
                    digits.push(chars[cursor]);
                    cursor += 1;
                }
                if cursor < chars.len() && chars[cursor] == ')' {
                    repeat = digits.parse::<usize>().unwrap_or(1);
                    index = cursor;
                }
            }
            count += repeat;
        }
        index += 1;
    }

    count
}

fn count_pic_markers_after_decimal(pic: &str, markers: &[char]) -> usize {
    let upper = pic.trim().to_ascii_uppercase();
    let chars: Vec<char> = upper.chars().collect();
    let Some(mut index) = chars.iter().position(|ch| *ch == 'V' || *ch == '.') else {
        return 0;
    };

    index += 1;
    let mut count = 0;

    while index < chars.len() {
        let ch = chars[index];
        if markers.contains(&ch) {
            let mut repeat = 1usize;
            if index + 1 < chars.len() && chars[index + 1] == '(' {
                let mut cursor = index + 2;
                let mut digits = String::new();
                while cursor < chars.len() && chars[cursor] != ')' {
                    digits.push(chars[cursor]);
                    cursor += 1;
                }
                if cursor < chars.len() && chars[cursor] == ')' {
                    repeat = digits.parse::<usize>().unwrap_or(1);
                    index = cursor;
                }
            }
            count += repeat;
        }
        index += 1;
    }

    count
}

// ════════════════════════════════════════════════════════════════════════════
// Procedure Division
// ════════════════════════════════════════════════════════════════════════════

fn walk_procedure_division(
    pair: Pair<Rule>,
    body: &mut Vec<Statement>,
    ctx: &CobolWalkerContext,
) -> Result<(), String> {
    let mut entry_name: Option<String> = None;
    let mut entry_params = Vec::new();
    let mut entry_return_type: Option<String> = None;
    let mut saw_leading_statements = false;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::using_clause => {
                entry_params = walk_using_params(child);
            }
            Rule::returning_clause => {
                entry_return_type = child
                    .into_inner()
                    .find(|part| part.as_rule() == Rule::ident_name)
                    .map(|part| part.as_str().to_string());
            }
            Rule::statement_list => {
                let mut statements = Vec::new();
                walk_statement_list(child, &mut statements, ctx)?;
                if !statements.is_empty() && entry_name.is_none() {
                    saw_leading_statements = true;
                }
                body.extend(statements);
            }
            Rule::declaratives_block => {
                // Parsed for compatibility; no shared AST lowering yet.
            }
            Rule::procedure_block => {
                for part in child.into_inner() {
                    match part.as_rule() {
                        Rule::section => {
                            if entry_name.is_none() && !saw_leading_statements {
                                entry_name = first_cobol_block_name(&part);
                            }
                            walk_section(part, body, ctx)?;
                        }
                        Rule::paragraph => {
                            if entry_name.is_none() && !saw_leading_statements {
                                entry_name = first_cobol_block_name(&part);
                            }
                            walk_paragraph(part, body, ctx)?;
                        }
                        _ => {}
                    }
                }
            }
            Rule::section => {
                if entry_name.is_none() && !saw_leading_statements {
                    entry_name = first_cobol_block_name(&child);
                }
                walk_section(child, body, ctx)?;
            }
            Rule::paragraph => {
                if entry_name.is_none() && !saw_leading_statements {
                    entry_name = first_cobol_block_name(&child);
                }
                walk_paragraph(child, body, ctx)?;
            }
            Rule::period => {}
            _ => {
                if !is_kw(child.as_rule()) {
                    // Skip unexpected tokens
                }
            }
        }
    }

    if let Some(name) = entry_name {
        if let Some(Statement {
            kind: StmtKind::FunctionDecl { params, return_type, is_sub, .. },
            ..
        }) = body.iter_mut().find(|stmt| {
            matches!(&stmt.kind, StmtKind::FunctionDecl { name: fn_name, .. } if fn_name.eq_ignore_ascii_case(&name))
        }) {
            *params = entry_params;
            *return_type = entry_return_type.clone();
            *is_sub = entry_return_type.is_none();
        }

        body.push(Statement::new(StmtKind::Expr(Expression::new(
            ExprKind::Call {
                callee: Box::new(Expression::ident(&name)),
                args: Vec::new(),
                optional: false },
        ))));
    }

    Ok(())
}

fn first_cobol_block_name(pair: &Pair<Rule>) -> Option<String> {
    pair.clone()
        .into_inner()
        .find(|child| child.as_rule() == Rule::paragraph_name)
        .map(|child| child.as_str().to_string())
}

fn walk_paragraph(
    pair: Pair<Rule>,
    body: &mut Vec<Statement>,
    ctx: &CobolWalkerContext,
) -> Result<(), String> {
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
                walk_statement_list(p, &mut para_body, ctx)?;
            }
            _ => {}
        }
    }

    if name.is_empty() {
        return Ok(());
    }

    let is_generator = body_has_yield(&para_body);

    body.push(Statement::with_span(
        StmtKind::FunctionDecl {
            name,
            params: Vec::new(),
            return_type: None,
            body: para_body,
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async: false,
            is_generator,
            is_sub: true },
        span,
    ));
    Ok(())
}

fn walk_section(
    pair: Pair<Rule>,
    body: &mut Vec<Statement>,
    ctx: &CobolWalkerContext,
) -> Result<(), String> {
    let span = to_span(&pair);
    let parts = filter_nokw(pair.into_inner());

    let mut name = String::new();
    let mut section_body = Vec::new();

    for p in parts {
        match p.as_rule() {
            Rule::paragraph_name => {
                name = p.as_str().to_string();
            }
            Rule::statement_list => {
                walk_statement_list(p, &mut section_body, ctx)?;
            }
            Rule::paragraph => {
                walk_paragraph(p, &mut section_body, ctx)?;
            }
            _ => {}
        }
    }

    if name.is_empty() {
        return Ok(());
    }

    let is_generator = body_has_yield(&section_body);

    body.push(Statement::with_span(
        StmtKind::FunctionDecl {
            name,
            params: Vec::new(),
            return_type: None,
            body: section_body,
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async: false,
            is_generator,
            is_sub: true },
        span,
    ));
    Ok(())
}

fn walk_statement_list(
    pair: Pair<Rule>,
    body: &mut Vec<Statement>,
    ctx: &CobolWalkerContext,
) -> Result<(), String> {
    for child in pair.into_inner() {
        let rule = child.as_rule();
        if matches!(rule, Rule::period) || is_kw(rule) {
            continue;
        }
        if let Some(stmt) = walk_statement(child, ctx)? {
            body.push(stmt);
        }
    }
    Ok(())
}

// ════════════════════════════════════════════════════════════════════════════
// Statements
// ════════════════════════════════════════════════════════════════════════════

fn walk_statement(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<Option<Statement>, String> {
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
            // DISPLAY of a SCREEN SECTION item renders to the terminal, not to
            // stdout — drop those operands. If every operand is a screen item,
            // the statement produces no output at all.
            exprs.retain(|e| !matches!(&e.kind, ExprKind::Ident(n) if ctx.is_screen_item(n)));
            if exprs.is_empty() {
                return Ok(None);
            }
            // Pad each field operand to the fixed width implied by its PICTURE
            // (numeric 9(n) → zero-pad, alphanumeric X(n) → space-pad). Literals
            // and computed expressions (e.g. FUNCTION results) are left untouched.
            for e in exprs.iter_mut() {
                if let ExprKind::Ident(n) = &e.kind {
                    if let Some(fmt) = ctx.field_pic(n) {
                        *e = cobol_pic_format_expr(std::mem::replace(e, Expression::null()), fmt);
                    }
                }
            }
            StmtKind::Echo(exprs)
        }

        // ── ACCEPT ──────────────────────────────────────────────────────
        Rule::accept_stmt => walk_accept_stmt(pair, ctx)?,

        // ── MOVE ────────────────────────────────────────────────────────
        Rule::move_stmt => walk_move_stmt(pair, ctx)?,

        // ── ADD ─────────────────────────────────────────────────────────
        Rule::add_stmt => walk_add_stmt(pair, ctx)?,

        // ── SUBTRACT ────────────────────────────────────────────────────
        Rule::subtract_stmt => walk_subtract_stmt(pair, ctx)?,

        // ── MULTIPLY ────────────────────────────────────────────────────
        Rule::multiply_stmt => walk_multiply_stmt(pair, ctx)?,

        // ── DIVIDE ──────────────────────────────────────────────────────
        Rule::divide_stmt => walk_divide_stmt(pair, ctx)?,

        // ── COMPUTE ─────────────────────────────────────────────────────
        Rule::compute_stmt => walk_compute_stmt(pair, ctx)?,

        // ── IF ──────────────────────────────────────────────────────────
        Rule::if_stmt => walk_if_stmt(pair, ctx)?,

        // ── EVALUATE ────────────────────────────────────────────────────
        Rule::evaluate_stmt => walk_evaluate_stmt(pair, ctx)?,

        // ── PERFORM ─────────────────────────────────────────────────────
        Rule::perform_stmt => walk_perform_stmt(pair, ctx)?,

        // ── STRING ──────────────────────────────────────────────────────
        Rule::string_stmt => walk_string_stmt(pair, ctx)?,

        // ── UNSTRING ────────────────────────────────────────────────────
        Rule::unstring_stmt => walk_unstring_stmt(pair, ctx)?,

        // ── INSPECT ─────────────────────────────────────────────────────
        Rule::inspect_stmt => walk_inspect_stmt(pair, ctx)?,

        // ── CALL ────────────────────────────────────────────────────────
        Rule::call_stmt => walk_call_stmt(pair)?,

        Rule::cancel_stmt => walk_cancel_stmt(pair)?,

        Rule::release_stmt => walk_release_stmt(pair)?,

        Rule::return_stmt => walk_return_stmt(pair, ctx)?,

        // ── INITIALIZE ──────────────────────────────────────────────────
        Rule::initialize_stmt => walk_initialize_stmt(pair)?,

        // ── SET ─────────────────────────────────────────────────────────
        Rule::set_stmt => walk_set_stmt(pair)?,

        // ── ALTER ──────────────────────────────────────────────────────
        Rule::alter_stmt => walk_alter_stmt(pair)?,

        // ── GO TO ───────────────────────────────────────────────────────
        Rule::go_to_stmt => {
            let parts = inner_nokw(pair);
            let name = parts
                .into_iter()
                .find(|p| matches!(p.as_rule(), Rule::ident_name | Rule::paragraph_name))
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            // GO TO paragraph → call the paragraph
            StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(&name)),
                args: Vec::new(),
                optional: false }))
        }

        // ── STOP RUN ────────────────────────────────────────────────────
        // Terminate the whole run via WASI cli/exit (halts from any depth,
        // including inside a PERFORMed paragraph), not a mere function return.
        Rule::stop_run_stmt => StmtKind::Expr(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__stop_run")),
            args: Vec::new(),
            optional: false })),

        Rule::stop_stmt => walk_stop_stmt(pair)?,

        // ── GOBACK ──────────────────────────────────────────────────────
        Rule::goback_stmt => StmtKind::Return(None),

        // ── CONTINUE ────────────────────────────────────────────────────
        Rule::continue_stmt => StmtKind::Empty,

        // ── RAISE ───────────────────────────────────────────────────────
        Rule::raise_stmt => {
            let parts = inner_nokw(pair);
            let expr = parts
                .into_iter()
                .find(|p| p.as_rule() == Rule::expression)
                .map(|p| walk_expression(p))
                .transpose()?;
            StmtKind::Throw { expr, cause: None }
        }

        // ── XML / JSON GENERATE / PARSE ─────────────────────────────────
        Rule::xml_stmt => walk_xml_stmt(pair)?,
        Rule::json_stmt => walk_json_stmt(pair)?,

        // ── OPEN ────────────────────────────────────────────────────────
        Rule::open_stmt => walk_open_stmt(pair, ctx)?,

        // ── CLOSE ───────────────────────────────────────────────────────
        Rule::close_stmt => walk_close_stmt(pair, ctx)?,

        // ── READ ────────────────────────────────────────────────────────
        Rule::read_stmt => walk_read_stmt(pair, ctx)?,

        // ── WRITE ───────────────────────────────────────────────────────
        Rule::write_stmt => walk_write_stmt(pair, ctx)?,

        // ── REWRITE ─────────────────────────────────────────────────────
        Rule::rewrite_stmt => walk_rewrite_stmt(pair, ctx)?,

        // ── DELETE ──────────────────────────────────────────────────────
        Rule::delete_stmt => walk_delete_stmt(pair, ctx)?,

        // ── START ───────────────────────────────────────────────────────
        Rule::start_stmt => walk_start_stmt(pair, ctx)?,

        // ── SORT ────────────────────────────────────────────────────────
        Rule::sort_stmt => walk_sort_stmt(pair)?,

        // ── MERGE ───────────────────────────────────────────────────────
        Rule::merge_stmt => walk_merge_stmt(pair)?,

        // ── SEARCH ──────────────────────────────────────────────────────
        Rule::search_stmt => walk_search_stmt(pair, ctx)?,

        // ── COPY ────────────────────────────────────────────────────────
        Rule::copy_stmt => walk_copy_stmt(pair)?,

        // ── INVOKE (OO COBOL) ──────────────────────────────────────────
        Rule::invoke_stmt => walk_invoke_stmt(pair)?,

        // ── VALIDATE ────────────────────────────────────────────────────
        Rule::validate_stmt => walk_validate_stmt(pair)?,

        // ── FREE ────────────────────────────────────────────────────────
        Rule::free_stmt => {
            let parts = inner_nokw(pair);
            let name = parts
                .into_iter()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            StmtKind::Assign {
                targets: vec![Expression::ident(&name)],
                value: Expression::null(), by_ref: false }
        }

        // ── ALLOCATE ────────────────────────────────────────────────────
        Rule::allocate_stmt => {
            let parts = inner_nokw(pair);
            let names: Vec<String> = parts
                .into_iter()
                .filter(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
                .collect();
            let target = names.first().cloned().unwrap_or_default();
            StmtKind::Assign {
                targets: vec![Expression::ident(&target)],
                value: Expression::new(ExprKind::Object(Vec::new())), by_ref: false }
        }

        // ── TYPEDEF ─────────────────────────────────────────────────────
        Rule::typedef_stmt => walk_typedef_stmt(pair)?,

        // ── EXIT ────────────────────────────────────────────────────────
        Rule::exit_stmt => walk_exit_stmt(pair)?,

        // ── WAIT ────────────────────────────────────────────────────────
        Rule::wait_stmt => {
            let parts = inner_nokw(pair);
            let name = parts
                .into_iter()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            StmtKind::Expr(Expression::new(ExprKind::Await(Box::new(
                Expression::ident(&name),
            ))))
        }

        // ── RUN UNIT ────────────────────────────────────────────────────
        Rule::run_unit_stmt => walk_run_unit_stmt(pair)?,

        // ── LOCK ────────────────────────────────────────────────────────
        Rule::lock_stmt => {
            let parts = inner_nokw(pair);
            let name = parts
                .into_iter()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__lock")),
                args: vec![Argument::positional(Expression::ident(&name))],
                optional: false }))
        }

        // ── UNLOCK ──────────────────────────────────────────────────────
        Rule::unlock_stmt => {
            let parts = inner_nokw(pair);
            let name = parts
                .into_iter()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__unlock")),
                args: vec![Argument::positional(Expression::ident(&name))],
                optional: false }))
        }

        // ── YIELD ───────────────────────────────────────────────────────
        Rule::yield_stmt => StmtKind::Expr(Expression::new(ExprKind::Yield(None))),

        // ── SUSPEND ─────────────────────────────────────────────────────
        Rule::suspend_stmt => StmtKind::Expr(Expression::new(ExprKind::Yield(None))),

        // ── EXEC SQL ────────────────────────────────────────────────────
        Rule::exec_sql_stmt => {
            let raw = extract_exec_body(&pair);
            StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__exec_sql")),
                args: vec![Argument::positional(Expression::string(&raw))],
                optional: false }))
        }

        // ── EXEC CICS ───────────────────────────────────────────────────
        Rule::exec_cics_stmt => {
            let raw = extract_exec_body(&pair);
            StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__exec_cics")),
                args: vec![Argument::positional(Expression::string(&raw))],
                optional: false }))
        }

        // ── EXEC DLI ────────────────────────────────────────────────────
        Rule::exec_dli_stmt => {
            let raw = extract_exec_body(&pair);
            StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__exec_dli")),
                args: vec![Argument::positional(Expression::string(&raw))],
                optional: false }))
        }

        // ── Nested program ──────────────────────────────────────────────
        Rule::nested_program_stmt => {
            // Compile inline — walk its procedure division
            let mut nested_body = Vec::new();
            let mut nested_ctx = CobolWalkerContext::new();
            for child in pair.into_inner() {
                match child.as_rule() {
                    Rule::procedure_division => {
                        walk_procedure_division(child, &mut nested_body, &nested_ctx)?;
                    }
                    Rule::data_division => {
                        walk_data_division(child, &mut nested_body, &mut nested_ctx)?;
                    }
                    _ => {}
                }
            }
            StmtKind::Block(nested_body)
        }

        // ── statement_list (transparent wrapper) ────────────────────────
        Rule::statement_list => {
            let mut stmts = Vec::new();
            walk_statement_list(pair, &mut stmts, ctx)?;
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
            return Err(format!(
                "COBOL walker: unhandled statement rule {:?}",
                other
            ));
        }
    };

    Ok(Some(Statement::with_span(kind, span)))
}

// ════════════════════════════════════════════════════════════════════════════
// Statement walkers
// ════════════════════════════════════════════════════════════════════════════

// ── ACCEPT ──────────────────────────────────────────────────────────────────

fn walk_accept_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // ACCEPT of a SCREEN SECTION item is interactive terminal input, which is
    // mocked away here — make it a no-op so it never issues a blocking read on
    // real stdin (which would hang an interactive test run). Mirrors the DISPLAY
    // suppression for screen items.
    if let Some(target) = children.iter().find(|c| c.as_rule() == Rule::ident_name) {
        if ctx.is_screen_item(target.as_str()) {
            return Ok(StmtKind::Block(Vec::new()));
        }
    }

    let mut var_name = String::new();
    let mut source: Option<String> = None;
    let mut args: Vec<Argument> = Vec::new();

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
                    match inner.as_rule() {
                        Rule::accept_environment_source => {
                            source = Some("ENVIRONMENT".to_string());
                            for env_inner in inner.into_inner() {
                                match env_inner.as_rule() {
                                    Rule::ident_name => {
                                        args.push(Argument::positional(Expression::ident(
                                            env_inner.as_str(),
                                        )));
                                    }
                                    Rule::string_literal => {
                                        args.push(Argument::positional(walk_string_literal(
                                            &env_inner,
                                        )?));
                                    }
                                    _ => {}
                                }
                            }
                        }
                        Rule::accept_date_source => {
                            source = Some(
                                if inner.as_str().to_ascii_uppercase().contains("YYYYMMDD") {
                                    "DATE-YYYYMMDD".to_string()
                                } else {
                                    "DATE".to_string()
                                },
                            );
                        }
                        Rule::accept_day_source => {
                            source =
                                Some(if inner.as_str().to_ascii_uppercase().contains("YYYYDDD") {
                                    "DAY-YYYYDDD".to_string()
                                } else {
                                    "DAY".to_string()
                                });
                        }
                        Rule::accept_day_of_week_source => {
                            source = Some("DAY-OF-WEEK".to_string());
                        }
                        Rule::kw_time => {
                            source = Some("TIME".to_string());
                        }
                        Rule::kw_command_line => {
                            source = Some("COMMAND-LINE".to_string());
                        }
                        Rule::kw_console => {
                            source = Some("CONSOLE".to_string());
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    let (call_callee, extra_args): (&str, Vec<Expression>) = match source.as_deref() {
        Some("ENVIRONMENT") => ("getenv", Vec::new()),
        Some("DATE") => ("__accept_date", vec![Expression::string("YYMMDD")]),
        Some("DATE-YYYYMMDD") => ("__accept_date", vec![Expression::string("YYYYMMDD")]),
        Some("DAY") => ("__accept_day", vec![Expression::string("YYDDD")]),
        Some("DAY-YYYYDDD") => ("__accept_day", vec![Expression::string("YYYYDDD")]),
        Some("DAY-OF-WEEK") => ("__accept_day_of_week", Vec::new()),
        Some("TIME") => ("__accept_time", Vec::new()),
        Some("COMMAND-LINE") => ("__accept_command_line", Vec::new()),
        Some("CONSOLE") => ("readline", Vec::new()),
        _ => ("readline", Vec::new()) };
    args.extend(extra_args.into_iter().map(Argument::positional));

    // ACCEPT var → var = readline()
    Ok(StmtKind::Assign {
        targets: vec![Expression::ident(&var_name)],
        value: Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(call_callee)),
            args,
            optional: false }), by_ref: false })
}

fn walk_cancel_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut target: Option<Expression> = None;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::string_literal => target = Some(walk_string_literal(&child)?),
            Rule::ident_name => target = Some(Expression::ident(child.as_str())),
            _ => {}
        }
    }

    Ok(StmtKind::Expr(cobol_helper_call(
        "__cobol_cancel",
        vec![target.unwrap_or_else(Expression::null)],
    )))
}

fn walk_release_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut record_name = String::new();
    let mut source_name: Option<String> = None;

    for child in pair.into_inner() {
        if child.as_rule() != Rule::ident_name {
            continue;
        }
        if record_name.is_empty() {
            record_name = child.as_str().to_string();
        } else if source_name.is_none() {
            source_name = Some(child.as_str().to_string());
        }
    }

    let source_expr = source_name
        .map(|name| Expression::ident(&name))
        .unwrap_or_else(|| Expression::ident(&record_name));
    Ok(StmtKind::Expr(cobol_helper_call(
        "__cobol_release",
        vec![Expression::string(&record_name), source_expr],
    )))
}

fn walk_return_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    let mut file_name = String::new();
    let mut into_var: Option<String> = None;
    let mut fail_body = Vec::new();
    let mut success_body = Vec::new();

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
                for clause in child.into_inner() {
                    if matches!(
                        clause.as_rule(),
                        Rule::statement_list | Rule::clause_statement_list
                    ) {
                        if fail_body.is_empty() {
                            walk_statement_list(clause, &mut fail_body, ctx)?;
                        } else {
                            walk_statement_list(clause, &mut success_body, ctx)?;
                        }
                    }
                }
            }
            _ => {}
        }
    }

    let return_call = cobol_helper_call("__cobol_return", vec![Expression::string(&file_name)]);
    let Some(target_name) = into_var else {
        return Ok(StmtKind::Expr(return_call));
    };

    let mut stmts = vec![Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(&target_name)],
        value: return_call, by_ref: false })];

    if let Some(status_stmt) = cobol_file_status_assign(
        ctx.file_binding(&file_name)
            .and_then(|binding| binding.status_var.as_deref()),
        "00",
    ) {
        success_body.insert(0, status_stmt);
    }

    if !fail_body.is_empty() || !success_body.is_empty() {
        stmts.push(Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Eq,
                Expression::ident(&target_name),
                Expression::null(),
            ),
            then_body: fail_body,
            elifs: Vec::new(),
            else_body: (!success_body.is_empty()).then_some(success_body) }));
    }

    Ok(if stmts.len() == 1 {
        stmts.remove(0).kind
    } else {
        StmtKind::Block(stmts)
    })
}

fn walk_stop_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut stmts = Vec::new();

    for child in pair.into_inner() {
        if child.as_rule() == Rule::literal {
            stmts.push(Statement::new(StmtKind::Echo(vec![walk_literal(child)?])));
        }
    }

    stmts.push(Statement::new(StmtKind::Return(None)));
    Ok(if stmts.len() == 1 {
        stmts.remove(0).kind
    } else {
        StmtKind::Block(stmts)
    })
}

fn walk_alter_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let names: Vec<String> = pair
        .into_inner()
        .filter(|child| child.as_rule() == Rule::paragraph_name)
        .map(|child| child.as_str().to_string())
        .collect();
    let source = names.first().cloned().unwrap_or_default();
    let target = names.get(1).cloned().unwrap_or_default();

    Ok(StmtKind::Expr(cobol_helper_call(
        "__cobol_alter",
        vec![Expression::string(&source), Expression::string(&target)],
    )))
}

fn walk_sort_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut file_name = String::new();
    let mut pending_order: Option<&'static str> = None;
    let mut keys = Vec::new();
    let mut using_operands = Vec::new();
    let mut giving_operands = Vec::new();
    let mut input_proc: Option<String> = None;
    let mut input_thru: Option<String> = None;
    let mut output_proc: Option<String> = None;
    let mut output_thru: Option<String> = None;
    let mut duplicates_in_order = false;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::ident_name if file_name.is_empty() => {
                file_name = child.as_str().to_string();
            }
            Rule::kw_ascending => pending_order = Some("ASCENDING"),
            Rule::kw_descending => pending_order = Some("DESCENDING"),
            Rule::ident_name => {
                if let Some(order) = pending_order.take() {
                    keys.push(cobol_object(vec![
                        ("name", Expression::string(child.as_str())),
                        ("order", Expression::string(order)),
                    ]));
                }
            }
            Rule::sort_duplicates_clause => {
                duplicates_in_order = true;
            }
            Rule::sort_input => {
                let (operands, proc_name, thru_name) = parse_cobol_sort_endpoint(child)?;
                using_operands.extend(operands);
                input_proc = proc_name;
                input_thru = thru_name;
            }
            Rule::sort_output => {
                let (operands, proc_name, thru_name) = parse_cobol_sort_endpoint(child)?;
                giving_operands.extend(operands);
                output_proc = proc_name;
                output_thru = thru_name;
            }
            _ => {}
        }
    }

    Ok(StmtKind::Expr(cobol_helper_call(
        "__cobol_sort",
        vec![cobol_object(vec![
            ("file", Expression::string(&file_name)),
            ("keys", cobol_array(keys)),
            ("duplicates_in_order", Expression::bool(duplicates_in_order)),
            ("using", cobol_array(using_operands)),
            ("giving", cobol_array(giving_operands)),
            ("input_procedure", option_string_expr(input_proc.as_deref())),
            ("input_thru", option_string_expr(input_thru.as_deref())),
            (
                "output_procedure",
                option_string_expr(output_proc.as_deref()),
            ),
            ("output_thru", option_string_expr(output_thru.as_deref())),
        ])],
    )))
}

fn walk_merge_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let merge_source = pair.as_str().to_ascii_uppercase();
    let children: Vec<Pair<Rule>> = pair.clone().into_inner().collect();
    let mut file_name = String::new();
    let mut pending_order: Option<&'static str> = None;
    let mut keys = Vec::new();
    let mut using_operands = Vec::new();
    let mut giving_operands = Vec::new();
    let mut output_proc: Option<String> = None;
    let mut duplicates_in_order = false;

    for child in children.iter().cloned() {
        match child.as_rule() {
            Rule::ident_name if file_name.is_empty() => {
                file_name = child.as_str().to_string();
            }
            Rule::kw_ascending => pending_order = Some("ASCENDING"),
            Rule::kw_descending => pending_order = Some("DESCENDING"),
            Rule::ident_name => {
                if let Some(order) = pending_order.take() {
                    keys.push(cobol_object(vec![
                        ("name", Expression::string(child.as_str())),
                        ("order", Expression::string(order)),
                    ]));
                }
            }
            Rule::sort_duplicates_clause => {
                duplicates_in_order = true;
            }
            Rule::string_literal => {
                using_operands.push(walk_string_literal(&child)?);
            }
            Rule::sort_file_operand => {
                using_operands.push(cobol_sort_operand_expr(child)?);
            }
            Rule::kw_giving => {}
            Rule::kw_output | Rule::kw_procedure | Rule::kw_is => {}
            _ => {
                if child.as_rule() == Rule::ident_name {
                    continue;
                }
                let text = child.as_str().to_ascii_uppercase();
                if text.contains("OUTPUT PROCEDURE") {
                    output_proc = extract_first_ident_name(&child);
                }
            }
        }
    }

    if merge_source.contains(" GIVING ") {
        for child in children.into_iter() {
            if child.as_rule() == Rule::sort_file_operand {
                giving_operands.push(cobol_sort_operand_expr(child)?);
            }
        }
    }

    Ok(StmtKind::Expr(cobol_helper_call(
        "__cobol_merge",
        vec![cobol_object(vec![
            ("file", Expression::string(&file_name)),
            ("keys", cobol_array(keys)),
            ("duplicates_in_order", Expression::bool(duplicates_in_order)),
            ("using", cobol_array(using_operands)),
            ("giving", cobol_array(giving_operands)),
            (
                "output_procedure",
                option_string_expr(output_proc.as_deref()),
            ),
        ])],
    )))
}

fn walk_copy_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut copybook = String::new();
    let mut library: Option<String> = None;
    let mut replacements = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::ident_name if copybook.is_empty() => {
                copybook = child.as_str().to_string();
            }
            Rule::ident_name if library.is_none() => {
                library = Some(child.as_str().to_string());
            }
            Rule::copy_replacement => {
                replacements.push(parse_cobol_copy_replacement(child)?);
            }
            _ => {}
        }
    }

    Ok(StmtKind::Expr(cobol_helper_call(
        "__cobol_copy",
        vec![cobol_object(vec![
            ("copybook", Expression::string(&copybook)),
            ("library", option_string_expr(library.as_deref())),
            ("replacements", cobol_array(replacements)),
        ])],
    )))
}

fn walk_validate_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let name = pair
        .into_inner()
        .find(|child| child.as_rule() == Rule::ident_name)
        .map(|child| child.as_str().to_string())
        .unwrap_or_default();
    Ok(StmtKind::Expr(cobol_helper_call(
        "__cobol_validate",
        vec![Expression::ident(&name)],
    )))
}

fn walk_typedef_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut type_name = String::new();
    let mut pic: Option<String> = None;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::ident_name if type_name.is_empty() => {
                type_name = child.as_str().to_string();
            }
            Rule::pic_clause => {
                pic = child
                    .into_inner()
                    .find(|part| part.as_rule() == Rule::pic_string)
                    .map(|part| part.as_str().to_string());
            }
            _ => {}
        }
    }

    Ok(StmtKind::Expr(cobol_helper_call(
        "__cobol_typedef",
        vec![
            Expression::string(&type_name),
            option_string_expr(pic.as_deref()),
        ],
    )))
}

fn cobol_helper_call(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false })
}

fn cobol_array(values: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Array(
        values
            .into_iter()
            .map(|value| ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false })
            .collect(),
    ))
}

fn cobol_object(entries: Vec<(&str, Expression)>) -> Expression {
    Expression::new(ExprKind::Object(
        entries
            .into_iter()
            .map(|(key, value)| ObjectProperty::KeyValue {
                key: Expression::string(key),
                value })
            .collect(),
    ))
}

fn option_string_expr(value: Option<&str>) -> Expression {
    value
        .map(Expression::string)
        .unwrap_or_else(Expression::null)
}

fn parse_cobol_sort_endpoint(
    pair: Pair<Rule>,
) -> Result<(Vec<Expression>, Option<String>, Option<String>), String> {
    let mut operands = Vec::new();
    let mut procedure_name: Option<String> = None;
    let mut thru_name: Option<String> = None;
    let mut saw_procedure = false;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::kw_input
            | Rule::kw_output
            | Rule::kw_procedure
            | Rule::kw_is
            | Rule::kw_using
            | Rule::kw_giving => {
                if matches!(child.as_rule(), Rule::kw_procedure) {
                    saw_procedure = true;
                }
            }
            Rule::sort_file_operand => {
                operands.push(cobol_sort_operand_expr(child)?);
            }
            Rule::ident_name => {
                if saw_procedure && procedure_name.is_none() {
                    procedure_name = Some(child.as_str().to_string());
                } else if saw_procedure && thru_name.is_none() {
                    thru_name = Some(child.as_str().to_string());
                }
            }
            Rule::string_literal => {
                operands.push(walk_string_literal(&child)?);
            }
            _ => {}
        }
    }

    Ok((operands, procedure_name, thru_name))
}

fn cobol_sort_operand_expr(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner_pairs = pair.clone().into_inner();
    let inner = inner_pairs.next().unwrap_or(pair);
    match inner.as_rule() {
        Rule::ident_name => Ok(Expression::string(inner.as_str())),
        Rule::string_literal => walk_string_literal(&inner),
        _ => Ok(Expression::string(inner.as_str())) }
}

fn parse_cobol_copy_replacement(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut values = Vec::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::pseudo_text | Rule::string_literal | Rule::ident_name => {
                values.push(match child.as_rule() {
                    Rule::string_literal => walk_string_literal(&child)?,
                    _ => Expression::string(child.as_str()) });
            }
            _ => {}
        }
    }

    Ok(cobol_object(vec![
        (
            "from",
            values.first().cloned().unwrap_or_else(Expression::null),
        ),
        (
            "to",
            values.get(1).cloned().unwrap_or_else(Expression::null),
        ),
    ]))
}

fn extract_first_ident_name(pair: &Pair<Rule>) -> Option<String> {
    pair.clone()
        .into_inner()
        .find(|child| child.as_rule() == Rule::ident_name)
        .map(|child| child.as_str().to_string())
}

// ── MOVE ────────────────────────────────────────────────────────────────────

fn cobol_member_expr(base_name: &str, field_name: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::ident(base_name)),
        field: field_name.to_string(),
        null_safe: false })
}

fn cobol_corresponding_fields(
    src_name: &str,
    dst_name: &str,
    ctx: &CobolWalkerContext,
    numeric_only: bool,
) -> Vec<String> {
    let src_fields = ctx.record_fields_for_name(src_name);
    let dst_fields = ctx.record_fields_for_name(dst_name);
    let mut fields = Vec::new();

    for dst_field in dst_fields {
        if numeric_only && !dst_field.numeric {
            continue;
        }
        if src_fields.iter().any(|src_field| {
            src_field.name.eq_ignore_ascii_case(&dst_field.name)
                && (!numeric_only || src_field.numeric)
        }) {
            fields.push(dst_field.name);
        }
    }

    fields
}

fn build_cobol_corresponding_stmt(
    src_name: &str,
    dst_name: &str,
    ctx: &CobolWalkerContext,
    op: Option<CompoundOp>,
) -> StmtKind {
    let fields = cobol_corresponding_fields(src_name, dst_name, ctx, op.is_some());
    if fields.is_empty() {
        return match op {
            Some(compound_op) => StmtKind::CompoundAssign {
                target: Expression::ident(dst_name),
                op: compound_op,
                value: Expression::ident(src_name) },
            None => StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__move_corresponding")),
                args: vec![
                    Argument::positional(Expression::ident(src_name)),
                    Argument::positional(Expression::ident(dst_name)),
                ],
                optional: false })) };
    }

    let mut stmts = Vec::new();
    for field_name in fields {
        let target = cobol_member_expr(dst_name, &field_name);
        let value = cobol_member_expr(src_name, &field_name);
        let stmt = match op {
            Some(compound_op) => StmtKind::CompoundAssign {
                target,
                op: compound_op,
                value },
            None => StmtKind::Assign {
                targets: vec![target],
                value, by_ref: false } };
        stmts.push(Statement::new(stmt));
    }

    if stmts.len() == 1 {
        stmts.remove(0).kind
    } else {
        StmtKind::Block(stmts)
    }
}

fn extract_cobol_rounded_mode(pair: Pair<Rule>) -> Option<String> {
    for child in pair.into_inner() {
        if child.as_rule() == Rule::rounded_mode_spec {
            if let Some(mode) = child.clone().into_inner().next() {
                return Some(mode.as_str().to_string());
            }
            let text = child.as_str().trim();
            if !text.is_empty() {
                return Some(text.to_string());
            }
        }
    }

    Some("DEFAULT".to_string())
}

fn parse_cobol_size_error_clause(
    pair: Pair<Rule>,
    ctx: &CobolWalkerContext,
) -> Result<(Vec<Statement>, Vec<Statement>), String> {
    let mut on_size_error = Vec::new();
    let mut not_on_size_error = Vec::new();
    let mut is_not_clause = false;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::kw_not => {
                is_not_clause = true;
            }
            Rule::statement_list | Rule::clause_statement_list => {
                let mut stmts = Vec::new();
                walk_statement_list(child, &mut stmts, ctx)?;
                if is_not_clause {
                    not_on_size_error = stmts;
                    is_not_clause = false;
                } else {
                    on_size_error = stmts;
                }
            }
            _ => {}
        }
    }

    Ok((on_size_error, not_on_size_error))
}

fn apply_cobol_rounding(expr: Expression, rounded_mode: Option<&str>) -> Expression {
    let Some(mode) = rounded_mode else {
        return expr;
    };

    let callee = match mode.to_ascii_uppercase().as_str() {
        "TRUNCATION" | "NEAREST-TOWARD-ZERO" => "f64_trunc",
        "TOWARD-GREATER" => "f64_ceil",
        "TOWARD-LESSER" => "f64_floor",
        "AWAY-FROM-ZERO" => "__cobol_round_away_from_zero",
        "PROHIBITED" => return expr,
        _ => "f64_nearest" };

    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(callee)),
        args: vec![Argument::positional(expr)],
        optional: false })
}

fn wrap_cobol_size_error(
    stmt: StmtKind,
    on_size_error: Vec<Statement>,
    not_on_size_error: Vec<Statement>,
) -> StmtKind {
    if on_size_error.is_empty() && not_on_size_error.is_empty() {
        return stmt;
    }

    let catches = if on_size_error.is_empty() {
        Vec::new()
    } else {
        vec![CatchClause {
            types: Vec::new(),
            var_name: Some("__cobol_size_error".into()),
            stack_var: None,
            body: on_size_error,
            when_clause: None }]
    };

    StmtKind::Try {
        body: vec![Statement::new(stmt)],
        catches,
        else_body: (!not_on_size_error.is_empty()).then_some(not_on_size_error),
        finally: None }
}

fn walk_move_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Check for MOVE CORRESPONDING
    let has_corr = children
        .iter()
        .any(|c| matches!(c.as_rule(), Rule::kw_corresponding | Rule::kw_corr));

    if has_corr {
        // MOVE CORRESPONDING src TO dst
        let idents: Vec<String> = children
            .iter()
            .filter(|c| c.as_rule() == Rule::ident_name)
            .map(|c| c.as_str().to_string())
            .collect();
        let src = idents.first().cloned().unwrap_or_default();
        let dst = idents.get(1).cloned().unwrap_or_default();

        Ok(build_cobol_corresponding_stmt(&src, &dst, ctx, None))
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
            if is_kw(child.as_rule()) {
                continue;
            }

            if !after_to {
                if child.as_rule() == Rule::expression {
                    src_expr = Some(walk_expression(child)?);
                }
            } else if let Some(target_expr) = walk_assignment_target_expr(child.clone())? {
                targets.push(target_expr);
            }
        }

        let value = src_expr.ok_or("MOVE missing source expression")?;
        Ok(StmtKind::Assign { targets, value , by_ref: false })
    }
}

fn extract_data_target_name(pair: Pair<Rule>) -> Option<String> {
    match pair.as_rule() {
        Rule::ident_name | Rule::ident_or_keyword | Rule::kw_sd => Some(pair.as_str().to_string()),
        Rule::data_target | Rule::move_target | Rule::giving_clause | Rule::remainder_clause => {
            pair.into_inner().find_map(extract_data_target_name)
        }
        _ => None }
}

fn walk_assignment_target_expr(pair: Pair<Rule>) -> Result<Option<Expression>, String> {
    match pair.as_rule() {
        Rule::data_target => Ok(Some(walk_data_target_expr(pair)?)),
        Rule::move_target | Rule::giving_clause | Rule::remainder_clause => {
            for child in pair.into_inner() {
                if let Some(target) = walk_assignment_target_expr(child)? {
                    return Ok(Some(target));
                }
            }
            Ok(None)
        }
        Rule::ident_name | Rule::ident_or_keyword | Rule::kw_sd => {
            Ok(Some(Expression::ident(pair.as_str())))
        }
        _ => Ok(None) }
}

// ── ADD ─────────────────────────────────────────────────────────────────────

fn walk_add_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Check for CORRESPONDING
    let has_corr = children
        .iter()
        .any(|c| matches!(c.as_rule(), Rule::kw_corresponding | Rule::kw_corr));

    let mut rounded_mode: Option<String> = None;
    let mut on_size_error = Vec::new();
    let mut not_on_size_error = Vec::new();

    for child in &children {
        match child.as_rule() {
            Rule::rounded_clause => {
                rounded_mode = extract_cobol_rounded_mode(child.clone());
            }
            Rule::size_error_clause => {
                let (on_body, not_on_body) = parse_cobol_size_error_clause(child.clone(), ctx)?;
                on_size_error = on_body;
                not_on_size_error = not_on_body;
            }
            _ => {}
        }
    }

    if has_corr {
        let idents: Vec<String> = children
            .iter()
            .filter(|c| c.as_rule() == Rule::ident_name)
            .map(|c| c.as_str().to_string())
            .collect();
        let src = idents.first().cloned().unwrap_or_default();
        let dst = idents.get(1).cloned().unwrap_or_default();
        return Ok(wrap_cobol_size_error(
            build_cobol_corresponding_stmt(&src, &dst, ctx, Some(CompoundOp::Add)),
            on_size_error,
            not_on_size_error,
        ));
    }

    // Collect expressions before TO, identifiers after TO
    let mut exprs: Vec<Expression> = Vec::new();
    let mut giving_name: Option<String> = None;
    let mut to_name: Option<String> = None;
    let mut in_giving = false;
    let mut in_to = false;

    for child in &children {
        match child.as_rule() {
            Rule::kw_to => {
                in_to = true;
                in_giving = false;
                continue;
            }
            Rule::kw_giving => {
                in_giving = true;
                in_to = false;
                continue;
            }
            _ => {}
        }
        if is_kw(child.as_rule())
            || matches!(
                child.as_rule(),
                Rule::size_error_clause | Rule::rounded_clause
            )
        {
            continue;
        }

        if in_giving {
            if let Some(target_name) = extract_data_target_name(child.clone()) {
                giving_name = Some(target_name);
            }
        } else if in_to {
            if let Some(target_name) = extract_data_target_name(child.clone()) {
                to_name = Some(target_name);
            } else if child.as_rule() == Rule::giving_clause {
                // giving_clause nested inside
                if let Some(target_name) = extract_data_target_name(child.clone()) {
                    giving_name = Some(target_name);
                }
            }
        } else if child.as_rule() == Rule::expression {
            exprs.push(walk_expression(child.clone())?);
        } else if child.as_rule() == Rule::arith_operand {
            if let Some(inner) = child.clone().into_inner().next() {
                exprs.push(walk_atom(inner)?);
            }
        } else if child.as_rule() == Rule::literal {
            exprs.push(walk_literal(child.clone())?);
        } else if child.as_rule() == Rule::giving_clause {
            if let Some(target_name) = extract_data_target_name(child.clone()) {
                giving_name = Some(target_name);
            }
        }
    }

    // Build the sum expression
    let sum_expr = build_sum_expr(&exprs);

    let stmt = if let Some(giving) = giving_name {
        // ADD a b GIVING c → c = a + b (+ to if present)
        let total = if let Some(ref to) = to_name {
            binary(BinOp::Add, sum_expr, Expression::ident(to))
        } else {
            sum_expr
        };
        StmtKind::Assign {
            targets: vec![Expression::ident(&giving)],
            value: apply_cobol_rounding(total, rounded_mode.as_deref()), by_ref: false }
    } else if let Some(to) = to_name {
        if rounded_mode.is_some() {
            StmtKind::Assign {
                targets: vec![Expression::ident(&to)],
                value: apply_cobol_rounding(
                    binary(BinOp::Add, Expression::ident(&to), sum_expr),
                    rounded_mode.as_deref(),
                ), by_ref: false }
        } else {
            StmtKind::CompoundAssign {
                target: Expression::ident(&to),
                op: CompoundOp::Add,
                value: sum_expr }
        }
    } else {
        StmtKind::Empty
    };

    Ok(wrap_cobol_size_error(
        stmt,
        on_size_error,
        not_on_size_error,
    ))
}

// ── SUBTRACT ────────────────────────────────────────────────────────────────

fn walk_subtract_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut rounded_mode: Option<String> = None;
    let mut on_size_error = Vec::new();
    let mut not_on_size_error = Vec::new();

    for child in &children {
        match child.as_rule() {
            Rule::rounded_clause => {
                rounded_mode = extract_cobol_rounded_mode(child.clone());
            }
            Rule::size_error_clause => {
                let (on_body, not_on_body) = parse_cobol_size_error_clause(child.clone(), ctx)?;
                on_size_error = on_body;
                not_on_size_error = not_on_body;
            }
            _ => {}
        }
    }

    let has_corr = children
        .iter()
        .any(|c| matches!(c.as_rule(), Rule::kw_corresponding | Rule::kw_corr));

    if has_corr {
        let idents: Vec<String> = children
            .iter()
            .filter(|c| c.as_rule() == Rule::ident_name)
            .map(|c| c.as_str().to_string())
            .collect();
        let src = idents.first().cloned().unwrap_or_default();
        let dst = idents.get(1).cloned().unwrap_or_default();
        return Ok(wrap_cobol_size_error(
            build_cobol_corresponding_stmt(&src, &dst, ctx, Some(CompoundOp::Sub)),
            on_size_error,
            not_on_size_error,
        ));
    }

    let mut src_expr: Option<Expression> = None;
    let mut from_name: Option<String> = None;
    let mut giving_name: Option<String> = None;
    let mut in_from = false;

    for child in &children {
        match child.as_rule() {
            Rule::kw_from => {
                in_from = true;
                continue;
            }
            _ => {}
        }
        if is_kw(child.as_rule())
            || matches!(
                child.as_rule(),
                Rule::size_error_clause | Rule::rounded_clause
            )
        {
            continue;
        }
        if child.as_rule() == Rule::giving_clause {
            if let Some(target_name) = extract_data_target_name(child.clone()) {
                giving_name = Some(target_name);
            }
            continue;
        }

        if in_from {
            if let Some(target_name) = extract_data_target_name(child.clone()) {
                from_name = Some(target_name);
            }
        } else if child.as_rule() == Rule::expression {
            src_expr = Some(walk_expression(child.clone())?);
        } else if child.as_rule() == Rule::arith_operand {
            if let Some(inner) = child.clone().into_inner().next() {
                src_expr = Some(walk_atom(inner)?);
            }
        } else if child.as_rule() == Rule::literal {
            src_expr = Some(walk_literal(child.clone())?);
        }
    }

    let src = src_expr.unwrap_or(Expression::int(0));

    let stmt = if let Some(giving) = giving_name {
        let from_expr = from_name
            .map(|n| Expression::ident(&n))
            .unwrap_or(Expression::int(0));
        StmtKind::Assign {
            targets: vec![Expression::ident(&giving)],
            value: apply_cobol_rounding(
                binary(BinOp::Sub, from_expr, src),
                rounded_mode.as_deref(),
            ), by_ref: false }
    } else if let Some(from) = from_name {
        if rounded_mode.is_some() {
            StmtKind::Assign {
                targets: vec![Expression::ident(&from)],
                value: apply_cobol_rounding(
                    binary(BinOp::Sub, Expression::ident(&from), src),
                    rounded_mode.as_deref(),
                ), by_ref: false }
        } else {
            StmtKind::CompoundAssign {
                target: Expression::ident(&from),
                op: CompoundOp::Sub,
                value: src }
        }
    } else {
        StmtKind::Empty
    };

    Ok(wrap_cobol_size_error(
        stmt,
        on_size_error,
        not_on_size_error,
    ))
}

// ── MULTIPLY ────────────────────────────────────────────────────────────────

fn walk_multiply_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut src_expr: Option<Expression> = None;
    let mut by_expr: Option<Expression> = None;
    let mut by_name: Option<String> = None;
    let mut giving_name: Option<String> = None;
    let mut rounded_mode: Option<String> = None;
    let mut on_size_error = Vec::new();
    let mut not_on_size_error = Vec::new();
    let mut in_by = false;

    for child in &children {
        match child.as_rule() {
            Rule::kw_by => {
                in_by = true;
                continue;
            }
            Rule::rounded_clause => {
                rounded_mode = extract_cobol_rounded_mode(child.clone());
                continue;
            }
            Rule::size_error_clause => {
                let (on_body, not_on_body) = parse_cobol_size_error_clause(child.clone(), ctx)?;
                on_size_error = on_body;
                not_on_size_error = not_on_body;
                continue;
            }
            _ => {}
        }
        if is_kw(child.as_rule())
            || matches!(
                child.as_rule(),
                Rule::size_error_clause | Rule::rounded_clause
            )
        {
            continue;
        }
        if child.as_rule() == Rule::giving_clause {
            if let Some(target_name) = extract_data_target_name(child.clone()) {
                giving_name = Some(target_name);
            }
            continue;
        }

        if in_by {
            if child.as_rule() == Rule::arith_operand {
                if let Some(inner) = child.clone().into_inner().next() {
                    by_expr = Some(walk_atom(inner)?);
                }
            } else if let Some(target_name) = extract_data_target_name(child.clone()) {
                by_name = Some(target_name.clone());
                by_expr = Some(Expression::ident(&target_name));
            }
        } else if child.as_rule() == Rule::arith_operand {
            if let Some(inner) = child.clone().into_inner().next() {
                src_expr = Some(walk_atom(inner)?);
            }
        } else if child.as_rule() == Rule::expression {
            src_expr = Some(walk_expression(child.clone())?);
        }
    }

    let src = src_expr.unwrap_or(Expression::int(1));

    let stmt = if let Some(giving) = giving_name {
        let by_expr = by_expr.unwrap_or_else(|| {
            by_name
                .map(|n| Expression::ident(&n))
                .unwrap_or(Expression::int(1))
        });
        StmtKind::Assign {
            targets: vec![Expression::ident(&giving)],
            value: apply_cobol_rounding(binary(BinOp::Mul, src, by_expr), rounded_mode.as_deref()), by_ref: false }
    } else if let Some(by) = by_name {
        if rounded_mode.is_some() {
            StmtKind::Assign {
                targets: vec![Expression::ident(&by)],
                value: apply_cobol_rounding(
                    binary(BinOp::Mul, Expression::ident(&by), src),
                    rounded_mode.as_deref(),
                ), by_ref: false }
        } else {
            StmtKind::CompoundAssign {
                target: Expression::ident(&by),
                op: CompoundOp::Mul,
                value: src }
        }
    } else {
        StmtKind::Empty
    };

    Ok(wrap_cobol_size_error(
        stmt,
        on_size_error,
        not_on_size_error,
    ))
}

// ── DIVIDE ──────────────────────────────────────────────────────────────────

fn walk_divide_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut exprs: Vec<Expression> = Vec::new();
    let mut giving_name: Option<String> = None;
    let mut remainder_name: Option<String> = None;
    let mut rounded_mode: Option<String> = None;
    let mut on_size_error = Vec::new();
    let mut not_on_size_error = Vec::new();
    let mut is_by = false;
    let mut is_into = false;

    for child in &children {
        match child.as_rule() {
            Rule::kw_by => {
                is_by = true;
                is_into = false;
                continue;
            }
            Rule::kw_into => {
                is_into = true;
                is_by = false;
                continue;
            }
            Rule::rounded_clause => {
                rounded_mode = extract_cobol_rounded_mode(child.clone());
                continue;
            }
            Rule::size_error_clause => {
                let (on_body, not_on_body) = parse_cobol_size_error_clause(child.clone(), ctx)?;
                on_size_error = on_body;
                not_on_size_error = not_on_body;
                continue;
            }
            _ => {}
        }
        if is_kw(child.as_rule())
            || matches!(
                child.as_rule(),
                Rule::size_error_clause | Rule::rounded_clause
            )
        {
            continue;
        }
        if child.as_rule() == Rule::remainder_clause {
            if let Some(target_name) = extract_data_target_name(child.clone()) {
                remainder_name = Some(target_name);
            }
            continue;
        }

        if child.as_rule() == Rule::expression {
            exprs.push(walk_expression(child.clone())?);
        } else if is_by || is_into {
            // After GIVING keyword
            if giving_name.is_none() {
                if let Some(target_name) = extract_data_target_name(child.clone()) {
                    giving_name = Some(target_name);
                }
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

    let stmt = if let Some(rem_name) = remainder_name {
        // Two assigns: c = a / b, r = a % b
        // Wrap in a block
        let div_assign = Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident(&target_name)],
            value: apply_cobol_rounding(
                binary(BinOp::IDiv, dividend.clone(), divisor.clone()),
                rounded_mode.as_deref(),
            ), by_ref: false });
        let rem_assign = Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident(&rem_name)],
            value: binary(BinOp::Mod, dividend, divisor), by_ref: false });
        StmtKind::Block(vec![div_assign, rem_assign])
    } else {
        StmtKind::Assign {
            targets: vec![Expression::ident(&target_name)],
            value: apply_cobol_rounding(
                binary(BinOp::Div, dividend, divisor),
                rounded_mode.as_deref(),
            ), by_ref: false }
    };

    Ok(wrap_cobol_size_error(
        stmt,
        on_size_error,
        not_on_size_error,
    ))
}

// ── COMPUTE ─────────────────────────────────────────────────────────────────

fn walk_compute_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let parts = inner_nokw(pair);
    let mut target: Option<Expression> = None;
    let mut expr: Option<Expression> = None;
    let mut rounded_mode: Option<String> = None;
    let mut on_size_error = Vec::new();
    let mut not_on_size_error = Vec::new();

    for p in parts {
        match p.as_rule() {
            Rule::data_target => {
                if target.is_none() {
                    target = Some(walk_data_target_expr(p)?);
                }
            }
            Rule::ident_name => {
                if target.is_none() {
                    target = Some(Expression::ident(p.as_str()));
                }
            }
            Rule::expression => {
                expr = Some(walk_expression(p)?);
            }
            Rule::rounded_clause => {
                rounded_mode = extract_cobol_rounded_mode(p);
            }
            Rule::size_error_clause => {
                let (on_body, not_on_body) = parse_cobol_size_error_clause(p, ctx)?;
                on_size_error = on_body;
                not_on_size_error = not_on_body;
            }
            _ => {}
        }
    }

    Ok(wrap_cobol_size_error(
        StmtKind::Assign {
            targets: vec![target.unwrap_or_else(|| Expression::ident(""))],
            value: apply_cobol_rounding(
                expr.unwrap_or(Expression::int(0)),
                rounded_mode.as_deref(),
            ), by_ref: false },
        on_size_error,
        not_on_size_error,
    ))
}

// ── IF ──────────────────────────────────────────────────────────────────────

fn walk_if_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut cond: Option<Expression> = None;
    let mut then_body = Vec::new();
    let mut else_body: Option<Vec<Statement>> = None;

    for child in children {
        match child.as_rule() {
            Rule::condition => {
                if cond.is_none() {
                    cond = Some(walk_condition(child, ctx)?);
                }
            }
            Rule::statement_list => {
                if cond.is_some() && then_body.is_empty() {
                    walk_statement_list(child, &mut then_body, ctx)?;
                }
            }
            Rule::else_clause => {
                let mut body = Vec::new();
                for ec in child.into_inner() {
                    if ec.as_rule() == Rule::statement_list {
                        walk_statement_list(ec, &mut body, ctx)?;
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
        else_body })
}

// ── EVALUATE ────────────────────────────────────────────────────────────────

fn walk_evaluate_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
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
                let case = walk_when_clause(child, ctx)?;
                cases.push(case);
            }
            Rule::when_other_clause => {
                let mut body = Vec::new();
                for wc in child.into_inner() {
                    if wc.as_rule() == Rule::statement_list {
                        walk_statement_list(wc, &mut body, ctx)?;
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
        default })
}

fn walk_when_clause(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<SwitchCase, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    let mut conditions = Vec::new();
    let mut body = Vec::new();

    for child in children {
        match child.as_rule() {
            Rule::when_value => {
                let val = walk_when_value(child, ctx)?;
                if let Some(cond) = val {
                    conditions.push(cond);
                }
            }
            Rule::statement_list => {
                walk_statement_list(child, &mut body, ctx)?;
            }
            _ => {}
        }
    }

    Ok(SwitchCase { conditions, body })
}

fn walk_when_value(
    pair: Pair<Rule>,
    ctx: &CobolWalkerContext,
) -> Result<Option<CaseCondition>, String> {
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
    let mut exprs = Vec::new();
    for child in &children {
        match child.as_rule() {
            Rule::expression => exprs.push(walk_expression(child.clone())?),
            Rule::condition => exprs.push(walk_condition(child.clone(), ctx)?),
            _ => {}
        }
    }

    if exprs.len() >= 2 {
        Ok(Some(CaseCondition::Range {
            from: exprs[0].clone(),
            to: exprs[1].clone() }))
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
        else_body: default })
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
            binary(
                BinOp::And,
                binary(BinOp::GtEq, from.clone(), from.clone()),
                binary(BinOp::LtEq, to.clone(), to.clone()),
            )
        }
        CaseCondition::Comparison { op: _, expr } => expr.clone() }
}

// ── PERFORM ─────────────────────────────────────────────────────────────────

fn walk_perform_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let inner = pair.into_inner().next().ok_or("empty PERFORM")?;

    match inner.as_rule() {
        Rule::perform_varying => walk_perform_varying(inner, ctx),
        Rule::perform_until => walk_perform_until(inner, ctx),
        Rule::perform_times => walk_perform_times(inner, ctx),
        Rule::perform_paragraph_until => walk_perform_paragraph_until(inner, ctx),
        Rule::perform_async => {
            let parts = inner_nokw(inner);
            let name = parts
                .into_iter()
                .find(|p| matches!(p.as_rule(), Rule::ident_name | Rule::paragraph_name))
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            Ok(StmtKind::Expr(Expression::new(ExprKind::Await(Box::new(
                Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(&name)),
                    args: Vec::new(),
                    optional: false }),
            )))))
        }
        Rule::perform_thru => {
            let parts = inner_nokw(inner);
            let names: Vec<String> = parts
                .into_iter()
                .filter(|p| matches!(p.as_rule(), Rule::ident_name | Rule::paragraph_name))
                .map(|p| p.as_str().to_string())
                .collect();
            // PERFORM para1 THRU para2 → call both paragraphs
            let mut stmts = Vec::new();
            for name in &names {
                stmts.push(Statement::new(StmtKind::Expr(Expression::new(
                    ExprKind::Call {
                        callee: Box::new(Expression::ident(name)),
                        args: Vec::new(),
                        optional: false },
                ))));
            }
            Ok(StmtKind::Block(stmts))
        }
        Rule::perform_paragraph => {
            let parts = inner_nokw(inner);
            let name = parts
                .into_iter()
                .find(|p| matches!(p.as_rule(), Rule::ident_name | Rule::paragraph_name))
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            Ok(StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(&name)),
                args: Vec::new(),
                optional: false })))
        }
        Rule::perform_inline => {
            let mut body = Vec::new();
            for child in inner.into_inner() {
                if child.as_rule() == Rule::statement_list {
                    walk_statement_list(child, &mut body, ctx)?;
                }
            }
            Ok(StmtKind::Block(body))
        }
        other => Err(format!(
            "COBOL walker: unhandled perform variant {:?}",
            other
        )) }
}

fn walk_perform_varying(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
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
            Rule::kw_from => {
                state = 1;
                continue;
            }
            Rule::kw_by => {
                state = 2;
                continue;
            }
            Rule::kw_until => {
                state = 3;
                continue;
            }
            Rule::expression => match state {
                1 => from_expr = Some(walk_expression(child)?),
                2 => by_expr = Some(walk_expression(child)?),
                _ => {}
            },
            Rule::condition => {
                until_cond = Some(walk_condition(child, ctx)?);
            }
            Rule::statement_list => {
                walk_statement_list(child, &mut body, ctx)?;
            }
            _ => {}
        }
    }

    let from = from_expr.unwrap_or(Expression::int(0));
    let by = by_expr.unwrap_or(Expression::int(1));
    // COBOL UNTIL = loop while NOT condition
    let cond = negate_expr(until_cond.unwrap_or(Expression::bool(false)));

    // Lower to a For loop so the BY increment lives in the update clause. This
    // way EXIT PERFORM CYCLE (Continue) still advances the loop variable —
    // otherwise the increment sits at the end of the body and Continue skips it,
    // hanging forever.
    let init = Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(&var_name)],
        value: from, by_ref: false });

    let update = Expression::new(ExprKind::Assign {
        target: Box::new(Expression::ident(&var_name)),
        value: Box::new(binary(BinOp::Add, Expression::ident(&var_name), by)) });

    Ok(StmtKind::For {
        init: Some(Box::new(init)),
        cond: Some(cond),
        update: Some(update),
        body })
}

fn walk_perform_until(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
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
                until_cond = Some(walk_condition(child, ctx)?);
            }
            Rule::statement_list => {
                walk_statement_list(child, &mut body, ctx)?;
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
            until: false })
    } else {
        Ok(StmtKind::While {
            cond,
            body,
            else_body: None })
    }
}

fn walk_perform_paragraph_until(
    pair: Pair<Rule>,
    ctx: &CobolWalkerContext,
) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut test_after = false;
    let mut until_cond: Option<Expression> = None;
    let mut paragraph_name: Option<String> = None;

    for child in children {
        match child.as_rule() {
            Rule::paragraph_name | Rule::ident_name => {
                paragraph_name = Some(child.as_str().to_string());
            }
            Rule::test_clause => {
                for tc in child.into_inner() {
                    if tc.as_rule() == Rule::kw_after {
                        test_after = true;
                    }
                }
            }
            Rule::condition => {
                until_cond = Some(walk_condition(child, ctx)?);
            }
            _ => {}
        }
    }

    let body = vec![Statement::new(StmtKind::Expr(Expression::new(
        ExprKind::Call {
            callee: Box::new(Expression::ident(&paragraph_name.unwrap_or_default())),
            args: Vec::new(),
            optional: false },
    )))];

    let cond = negate_expr(until_cond.unwrap_or(Expression::bool(false)));

    if test_after {
        Ok(StmtKind::DoWhile {
            body,
            cond,
            until: false })
    } else {
        Ok(StmtKind::While {
            cond,
            body,
            else_body: None })
    }
}

fn walk_perform_times(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
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
                walk_statement_list(child, &mut body, ctx)?;
            }
            _ => {}
        }
    }

    let n = count_expr.unwrap_or(Expression::int(0));
    let counter = "__i";

    let init = Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(counter)],
        value: Expression::int(0), by_ref: false });
    let cond = binary(BinOp::Lt, Expression::ident(counter), n);
    let update = Expression::new(ExprKind::Assign {
        target: Box::new(Expression::ident(counter)),
        value: Box::new(binary(
            BinOp::Add,
            Expression::ident(counter),
            Expression::int(1),
        )) });

    Ok(StmtKind::For {
        init: Some(Box::new(init)),
        cond: Some(cond),
        update: Some(update),
        body })
}

// ── STRING ──────────────────────────────────────────────────────────────────

fn walk_string_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut source_exprs: Vec<Expression> = Vec::new();
    let mut into_name = String::new();
    let mut overflow_body: Option<Vec<Statement>> = None;
    let mut not_overflow_body: Option<Vec<Statement>> = None;
    let mut ptr_name: Option<String> = None;

    for child in children {
        match child.as_rule() {
            Rule::string_source => {
                // A source is `expr DELIMITED BY (SIZE | delim)`.
                //  • DELIMITED BY SIZE  → the whole sending field, i.e. padded to
                //    its PICTURE width.
                //  • DELIMITED BY delim → characters up to the first delimiter.
                let mut src_expr: Option<Expression> = None;
                let mut delim_expr: Option<Expression> = None;
                let mut by_size = false;
                for sc in child.into_inner() {
                    match sc.as_rule() {
                        Rule::kw_size => by_size = true,
                        Rule::expression => {
                            if src_expr.is_none() {
                                src_expr = Some(walk_expression(sc)?);
                            } else {
                                delim_expr = Some(walk_expression(sc)?);
                            }
                        }
                        _ => {}
                    }
                }
                let Some(mut e) = src_expr else { continue };
                if by_size {
                    // DELIMITED BY SIZE sends the WHOLE field, i.e. its
                    // PICTURE representation — the same rendering DISPLAY
                    // uses. Only the alphanumeric half was implemented, so a
                    // numeric field arrived as its bare value: cobc produced
                    // `lit 005 abcd` where Vybe gave `lit 5 abcd`.
                    // `cobol_pic_format_expr` is exactly what DISPLAY applies.
                    if let ExprKind::Ident(n) = &e.kind
                        && let Some(fmt) = ctx.field_pic(n)
                    {
                        e = cobol_pic_format_expr(e, fmt);
                    }
                } else if let Some(d) = delim_expr {
                    // take chars up to the first delimiter: split(d)[0]
                    e = Expression::new(ExprKind::Index {
                        object: Box::new(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(e),
                                field: "split".to_string(),
                                null_safe: false })),
                            args: vec![Argument::positional(d)],
                            optional: false })),
                        index: Box::new(Expression::int(0)),
                        null_safe: false });
                }
                source_exprs.push(e);
            }
            Rule::ident_name => {
                // The INTO target
                into_name = child.as_str().to_string();
            }
            Rule::pointer_clause => {
                ptr_name = child
                    .clone()
                    .into_inner()
                    .find(|p| p.as_rule() == Rule::ident_name)
                    .map(|p| p.as_str().to_string());
            }
            Rule::overflow_clause => {
                let mut lists = child
                    .into_inner()
                    .filter(|p| p.as_rule() == Rule::statement_list);
                if let Some(l) = lists.next() {
                    let mut b = Vec::new();
                    walk_statement_list(l, &mut b, ctx)?;
                    overflow_body = Some(b);
                }
                if let Some(l) = lists.next() {
                    let mut b = Vec::new();
                    walk_statement_list(l, &mut b, ctx)?;
                    not_overflow_body = Some(b);
                }
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

    // WITH POINTER p: write starting at 1-based position p, then advance p by the
    // number of characters transferred (p is a plain numeric counter).
    if let Some(ptr) = ptr_name {
        let val = Expression::ident("__str_val");
        let prefix = Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__refmod")),
            args: vec![
                Argument::positional(Expression::ident(&into_name)),
                Argument::positional(Expression::int(0)),
                Argument::positional(binary(
                    BinOp::Sub,
                    Expression::ident(&ptr),
                    Expression::int(1),
                )),
            ],
            optional: false });
        let len = Expression::new(ExprKind::Member {
            object: Box::new(val.clone()),
            field: "length".to_string(),
            null_safe: false });
        return Ok(StmtKind::Block(vec![
            Statement::new(StmtKind::VarDecl {
                kind: VarDeclKind::Dim,
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident("__str_val".to_string()),
                    type_hint: None,
                    init: Some(concat_expr),
                    array_bounds: None,
                    with_events: false }] }),
            Statement::new(StmtKind::Assign {
                targets: vec![Expression::ident(&into_name)],
                value: binary(BinOp::Concat, prefix, val), by_ref: false }),
            Statement::new(StmtKind::Assign {
                targets: vec![Expression::ident(&ptr)],
                value: binary(BinOp::Add, Expression::ident(&ptr), len), by_ref: false }),
        ]));
    }

    let assign = Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(&into_name)],
        value: concat_expr, by_ref: false });

    // ON OVERFLOW fires when the assembled string is longer than the receiving
    // field's PICTURE width — the classic "doesn't fit" condition.
    if overflow_body.is_some() || not_overflow_body.is_some() {
        if let Some(CobolPicFmt::Alpha(width)) = ctx.field_pic(&into_name) {
            let len = Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident(&into_name)),
                field: "length".to_string(),
                null_safe: false });
            let cond = binary(BinOp::Gt, len, Expression::int(width as i64));
            return Ok(StmtKind::Block(vec![
                assign,
                Statement::new(StmtKind::If {
                    cond,
                    then_body: overflow_body.unwrap_or_default(),
                    elifs: Vec::new(),
                    else_body: not_overflow_body }),
            ]));
        }
    }

    Ok(assign.kind)
}

// ── UNSTRING ────────────────────────────────────────────────────────────────

fn walk_unstring_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut src_name = String::new();
    let mut target_exprs: Vec<Expression> = Vec::new();
    let mut target_names: Vec<Option<String>> = Vec::new();
    let mut count_vars: Vec<Option<String>> = Vec::new();
    let mut delimiter: Option<Expression> = None;
    let mut saw_delimited_by = false;
    let mut delimited_by_all = false;
    let mut tally_var: Option<String> = None;

    for child in &children {
        match child.as_rule() {
            Rule::tallying_in_clause => {
                tally_var = child
                    .clone()
                    .into_inner()
                    .find(|p| p.as_rule() == Rule::ident_name)
                    .map(|p| p.as_str().to_string());
            }
            Rule::ident_name => {
                if src_name.is_empty() {
                    src_name = child.as_str().to_string();
                } else if saw_delimited_by && delimiter.is_none() {
                    delimiter = Some(Expression::ident(child.as_str()));
                }
            }
            Rule::kw_delimited | Rule::kw_by => {
                saw_delimited_by = true;
            }
            // DELIMITED BY ALL "x": consecutive delimiters collapse into one,
            // so empty tokens between them are dropped after the split.
            Rule::kw_all => {
                delimited_by_all = true;
            }
            Rule::string_literal => {
                if saw_delimited_by && delimiter.is_none() {
                    delimiter = Some(walk_string_literal(child)?);
                }
            }
            Rule::figurative_constant => {
                if saw_delimited_by && delimiter.is_none() {
                    delimiter = Some(walk_figurative_constant(child.clone())?);
                }
            }
            Rule::unstring_target => {
                // receiver [DELIMITER IN d] [COUNT IN c]
                let mut receiver: Option<Pair<Rule>> = None;
                let mut count_name: Option<String> = None;
                let mut state = 0u8; // 0=receiver, 1=after DELIMITER, 2=after COUNT
                for ut in child.clone().into_inner() {
                    match ut.as_rule() {
                        Rule::kw_delimiter => state = 1,
                        Rule::kw_count => state = 2,
                        Rule::data_target => {
                            if state == 0 && receiver.is_none() {
                                receiver = Some(ut.clone());
                            } else if state == 2 {
                                count_name = extract_data_target_name(ut.clone());
                            }
                        }
                        _ => {}
                    }
                }
                if let Some(r) = receiver {
                    target_names.push(extract_data_target_name(r.clone()));
                    target_exprs.push(walk_data_target_expr(r)?);
                    count_vars.push(count_name);
                }
            }
            _ => {}
        }
    }

    // The whole sending field participates in UNSTRING, so pad the source to its
    // PICTURE width first — that makes token lengths (COUNT IN) match COBOL.
    let src_expr = match ctx.field_pic(&src_name) {
        Some(CobolPicFmt::Alpha(w)) => Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__pad_end")),
            args: vec![
                Argument::positional(Expression::ident(&src_name)),
                Argument::positional(Expression::int(w as i64)),
                Argument::positional(Expression::string(" ")),
            ],
            optional: false }),
        _ => Expression::ident(&src_name) };
    let split_call = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(src_expr),
            field: "split".to_string(),
            null_safe: false })),
        args: vec![Argument::positional(
            delimiter.unwrap_or(Expression::string(" ")),
        )],
        optional: false });

    // DELIMITED BY ALL collapses runs of the delimiter, so the empty tokens the
    // plain split leaves between adjacent delimiters are dropped via .filter —
    // routed through the shared HOF loop emitter ([array_methods] in profile).
    let source_tokens = if delimited_by_all {
        let predicate = Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: "__t".to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false }],
            body: LambdaBody::Expr(Box::new(binary(
                BinOp::StrictNotEq,
                Expression::ident("__t"),
                Expression::string(""),
            ))),
            is_async: false,
            captures: vec![] });
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(split_call),
                field: "filter".to_string(),
                null_safe: false })),
            args: vec![Argument::positional(predicate)],
            optional: false })
    } else {
        split_call
    };

    let mut stmts = Vec::new();
    stmts.push(Statement::new(StmtKind::VarDecl {
        kind: VarDeclKind::Dim,
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident("__split_result".to_string()),
            type_hint: None,
            init: Some(source_tokens),
            array_bounds: None,
            with_events: false }] }));

    for (i, target) in target_exprs.iter().enumerate() {
        let token = Expression::new(ExprKind::Index {
            object: Box::new(Expression::ident("__split_result")),
            index: Box::new(Expression::int(i as i64)),
            null_safe: false });
        // Receiver gets the token truncated to its field width (left-justified).
        let value = match target_names
            .get(i)
            .and_then(|n| n.as_ref())
            .and_then(|n| ctx.field_pic(n))
        {
            Some(CobolPicFmt::Alpha(w)) => Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__refmod")),
                args: vec![
                    Argument::positional(token.clone()),
                    // __refmod takes a 0-based start (the refmod path already
                    // subtracts 1); truncate to the receiver width from index 0.
                    Argument::positional(Expression::int(0)),
                    Argument::positional(Expression::int(w as i64)),
                ],
                optional: false }),
            _ => token.clone() };
        stmts.push(Statement::new(StmtKind::Assign {
            targets: vec![target.clone()],
            value, by_ref: false }));
        // COUNT IN = the token's actual length.
        if let Some(Some(cv)) = count_vars.get(i) {
            stmts.push(Statement::new(StmtKind::Assign {
                targets: vec![Expression::ident(cv)],
                value: Expression::new(ExprKind::Member {
                    object: Box::new(token),
                    field: "length".to_string(),
                    null_safe: false }), by_ref: false }));
        }
    }

    // TALLYING IN counter = number of receivers actually filled =
    // min(token count, receiver count).
    if let Some(tv) = tally_var {
        let n = target_exprs.len() as i64;
        let split_len = Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident("__split_result")),
            field: "length".to_string(),
            null_safe: false });
        let count = Expression::new(ExprKind::Ternary {
            cond: Box::new(binary(BinOp::Gt, split_len.clone(), Expression::int(n))),
            then: Box::new(Expression::int(n)),
            else_: Box::new(split_len) });
        stmts.push(Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident(&tv)],
            value: count, by_ref: false }));
    }

    Ok(StmtKind::Block(stmts))
}

// ── INSPECT ─────────────────────────────────────────────────────────────────

fn walk_inspect_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let inner = pair.into_inner().next().ok_or("empty INSPECT")?;

    match inner.as_rule() {
        Rule::inspect_tallying => walk_inspect_tallying(inner),
        Rule::inspect_replacing => walk_inspect_replacing(inner, ctx),
        Rule::inspect_converting => walk_inspect_converting(inner),
        other => Err(format!(
            "COBOL walker: unhandled inspect variant {:?}",
            other
        )) }
}

fn walk_inspect_tallying(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let parts = inner_nokw(pair);
    let idents: Vec<String> = parts
        .iter()
        .filter(|p| p.as_rule() == Rule::ident_name)
        .map(|p| p.as_str().to_string())
        .collect();

    let var = idents.first().cloned().unwrap_or_default();
    let counter = idents.get(1).cloned().unwrap_or_default();

    // Find the target string/ident in tally phrases
    let mut target_expr = Expression::string(" ");
    let mut count_characters = false;
    for p in parts {
        if p.as_rule() == Rule::inspect_tally_phrase {
            for tp in p.into_inner() {
                if tp.as_rule() == Rule::kw_characters {
                    count_characters = true;
                    break;
                }
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

    if count_characters {
        let len_expr = Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&var)),
            field: "length".to_string(),
            null_safe: false });

        return Ok(StmtKind::Assign {
            targets: vec![Expression::ident(&counter)],
            value: len_expr, by_ref: false });
    }

    // counter = split(var, target).length - 1
    let split_call = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&var)),
            field: "split".to_string(),
            null_safe: false })),
        args: vec![Argument::positional(target_expr)],
        optional: false });

    let len_expr = Expression::new(ExprKind::Member {
        object: Box::new(split_call),
        field: "length".to_string(),
        null_safe: false });

    Ok(StmtKind::Assign {
        targets: vec![Expression::ident(&counter)],
        value: binary(BinOp::Sub, len_expr, Expression::int(1)), by_ref: false })
}

fn walk_inspect_replacing(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let parts = inner_nokw(pair);
    let var = parts
        .iter()
        .find(|p| p.as_rule() == Rule::ident_name)
        .map(|p| p.as_str().to_string())
        .unwrap_or_default();

    let mut old_expr = Expression::string("");
    let mut new_expr = Expression::string("");

    for p in &parts {
        if p.as_rule() == Rule::inspect_replace_phrase {
            // REPLACING CHARACTERS BY x — no search operand; every position in
            // the receiver becomes x, i.e. x cycled to the field width. `padEnd`
            // of the empty string fills the whole PICTURE width with the fill.
            if p.clone()
                .into_inner()
                .next()
                .is_some_and(|c| c.as_rule() == Rule::kw_characters)
            {
                let fill = p
                    .clone()
                    .into_inner()
                    .filter(|rp| {
                        rp.as_rule() == Rule::string_literal || rp.as_rule() == Rule::ident_name
                    })
                    .last();
                let fill_expr = match fill {
                    Some(rp) if rp.as_rule() == Rule::string_literal => walk_string_literal(&rp)?,
                    Some(rp) => Expression::ident(rp.as_str()),
                    None => Expression::string(" ") };
                let width = match ctx.field_pic(&var) {
                    Some(CobolPicFmt::Alpha(w)) | Some(CobolPicFmt::Numeric(w)) => w as i64,
                    None => 0 };
                let filled = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__pad_end")),
                    args: vec![
                        Argument::positional(Expression::string("")),
                        Argument::positional(Expression::int(width)),
                        Argument::positional(fill_expr),
                    ],
                    optional: false });
                return Ok(StmtKind::Assign {
                    targets: vec![Expression::ident(&var)],
                    value: filled, by_ref: false });
            }
            let mut found_by = false;
            for rp in p.clone().into_inner() {
                if rp.as_rule() == Rule::kw_by {
                    found_by = true;
                    continue;
                }
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
            null_safe: false })),
        args: vec![
            Argument::positional(old_expr),
            Argument::positional(new_expr),
        ],
        optional: false });

    Ok(StmtKind::Assign {
        targets: vec![Expression::ident(&var)],
        value: replace_call, by_ref: false })
}

fn walk_inspect_converting(pair: Pair<Rule>) -> Result<StmtKind, String> {
    // Keep raw inner so the `TO` keyword survives — it separates the from/to
    // character sets (inner_nokw strips it, collapsing both into `from`).
    let parts: Vec<Pair<Rule>> = pair.into_inner().collect();
    let var = parts
        .iter()
        .find(|p| p.as_rule() == Rule::ident_name)
        .map(|p| p.as_str().to_string())
        .unwrap_or_default();

    // INSPECT var CONVERTING from TO to → character replacement
    let mut from_expr = Expression::string("");
    let mut to_expr = Expression::string("");
    let mut found_to = false;

    for p in &parts {
        match p.as_rule() {
            Rule::kw_to => {
                found_to = true;
            }
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

    // CONVERTING is a per-character translation. The overwhelmingly common case
    // (and what has a working cobol path) is case folding: map the alphabet to
    // UPPER-CASE / LOWER-CASE, which route to the shared str_to_upper/lower.
    // Arbitrary translations fall back to a substring replace (imperfect).
    const LOWER: &str = "abcdefghijklmnopqrstuvwxyz";
    const UPPER: &str = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    let str_lit = |e: &Expression| match &e.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.clone()),
        _ => None };
    let value = match (str_lit(&from_expr), str_lit(&to_expr)) {
        (Some(f), Some(t)) if f == LOWER && t == UPPER => Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("UPPER-CASE")),
            args: vec![Argument::positional(Expression::ident(&var))],
            optional: false }),
        (Some(f), Some(t)) if f == UPPER && t == LOWER => Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("LOWER-CASE")),
            args: vec![Argument::positional(Expression::ident(&var))],
            optional: false }),
        _ => Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident(&var)),
                field: "replace".to_string(),
                null_safe: false })),
            args: vec![
                Argument::positional(from_expr),
                Argument::positional(to_expr),
            ],
            optional: false }) };

    Ok(StmtKind::Assign {
        targets: vec![Expression::ident(&var)],
        value, by_ref: false })
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
                    callee_name = s[1..s.len() - 1].to_string();
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
        optional: false });

    if let Some(ret_var) = returning_var {
        Ok(StmtKind::Assign {
            targets: vec![Expression::ident(&ret_var)],
            value: call_expr, by_ref: false })
    } else {
        Ok(StmtKind::Expr(call_expr))
    }
}

// ── INITIALIZE ──────────────────────────────────────────────────────────────

fn walk_initialize_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let parts = inner_nokw(pair);
    let names: Vec<String> = parts
        .into_iter()
        .filter(|p| p.as_rule() == Rule::ident_name)
        .map(|p| p.as_str().to_string())
        .collect();

    // INITIALIZE sets to default value (spaces for alpha, zeros for numeric)
    // Without knowing the PIC at walk time, default to empty string
    let mut stmts = Vec::new();
    for name in names {
        stmts.push(Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident(&name)],
            value: Expression::string(""), by_ref: false }));
    }

    Ok(if stmts.len() == 1 {
        stmts.remove(0).kind
    } else {
        StmtKind::Block(stmts)
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
            Rule::kw_true => {
                value = Some(Expression::bool(true));
            }
            Rule::kw_false => {
                value = Some(Expression::bool(false));
            }
            Rule::kw_up => {
                is_up = true;
            }
            Rule::kw_down => {
                is_down = true;
            }
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
            value: value.unwrap_or(Expression::int(1)) })
    } else if is_down {
        // SET var DOWN BY n → var -= n
        Ok(StmtKind::CompoundAssign {
            target: Expression::ident(&target),
            op: CompoundOp::Sub,
            value: value.unwrap_or(Expression::int(1)) })
    } else {
        Ok(StmtKind::Assign {
            targets: vec![Expression::ident(&target)],
            value: value.unwrap_or(Expression::bool(true)), by_ref: false })
    }
}

// ── JSON ────────────────────────────────────────────────────────────────────

fn walk_xml_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    let has_generate = children.iter().any(|c| c.as_rule() == Rule::kw_generate);
    let has_parse = children.iter().any(|c| c.as_rule() == Rule::kw_parse);

    if has_generate {
        let mut target = String::new();
        let mut source: Option<Expression> = None;
        let mut saw_from = false;

        for child in children {
            match child.as_rule() {
                Rule::kw_from => saw_from = true,
                Rule::ident_name if !saw_from && target.is_empty() => {
                    target = child.as_str().to_string();
                }
                Rule::expression if saw_from && source.is_none() => {
                    source = Some(walk_expression(child)?);
                }
                _ => {}
            }
        }

        Ok(StmtKind::Assign {
            targets: vec![Expression::ident(&target)],
            value: Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("xml_generate")),
                args: vec![Argument::positional(source.unwrap_or(Expression::null()))],
                optional: false }), by_ref: false })
    } else if has_parse {
        let mut source: Option<Expression> = None;
        let mut handler = String::new();
        let mut saw_processing = false;

        for child in children {
            match child.as_rule() {
                Rule::expression if source.is_none() => {
                    source = Some(walk_expression(child)?);
                }
                Rule::kw_processing => saw_processing = true,
                Rule::ident_name if saw_processing && handler.is_empty() => {
                    handler = child.as_str().to_string();
                }
                _ => {}
            }
        }

        Ok(StmtKind::Expr(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("xml_parse")),
            args: vec![
                Argument::positional(source.unwrap_or(Expression::null())),
                Argument::positional(Expression::ident(&handler)),
            ],
            optional: false })))
    } else {
        Ok(StmtKind::Empty)
    }
}

fn walk_json_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let has_generate = children.iter().any(|c| c.as_rule() == Rule::kw_generate);
    let has_parse = children.iter().any(|c| c.as_rule() == Rule::kw_parse);

    let idents: Vec<String> = children
        .iter()
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
                optional: false }), by_ref: false })
    } else if has_parse {
        // JSON PARSE src INTO dst → dst = json_decode(src)
        let src = idents.first().cloned().unwrap_or_default();
        let dst = idents.get(1).cloned().unwrap_or_default();
        Ok(StmtKind::Assign {
            targets: vec![Expression::ident(&dst)],
            value: Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("json_decode")),
                args: vec![Argument::positional(Expression::ident(&src))],
                optional: false }), by_ref: false })
    } else {
        Ok(StmtKind::Empty)
    }
}

// ── File I/O ────────────────────────────────────────────────────────────────

fn walk_open_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    let mut stmts = Vec::new();
    let mut current_mode = FileMode::Input;

    for child in children {
        match child.as_rule() {
            Rule::file_open_mode => {
                let mode_text = child.as_str().to_uppercase();
                current_mode = if mode_text.contains("INPUT") {
                    FileMode::Input
                } else if mode_text.contains("OUTPUT") {
                    FileMode::Output
                } else if mode_text.contains("EXTEND") {
                    FileMode::Append
                } else {
                    FileMode::Input
                };
            }
            Rule::ident_name => {
                let file_name = child.as_str().to_string();
                let binding =
                    ctx.file_binding(&file_name)
                        .cloned()
                        .unwrap_or_else(|| CobolFileBinding {
                            path: Expression::string(&file_name.to_ascii_lowercase()),
                            file_number: 1,
                            status_var: None,
                            key_name: None,
                            record_name: None,
                            record_fields: Vec::new() });
                stmts.push(Statement::new(StmtKind::OpenFile {
                    path: binding.path,
                    mode: current_mode,
                    file_number: Expression::int(binding.file_number as i64) }));
                if let Some(status_stmt) =
                    cobol_file_status_assign(binding.status_var.as_deref(), "00")
                {
                    stmts.push(status_stmt);
                }
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

fn walk_close_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let parts = inner_nokw(pair);
    let names: Vec<String> = parts
        .into_iter()
        .filter(|p| p.as_rule() == Rule::ident_name)
        .map(|p| p.as_str().to_string())
        .collect();

    let mut stmts = Vec::new();
    for name in names {
        let file_number = ctx
            .file_binding(&name)
            .map(|binding| Expression::int(binding.file_number as i64))
            .unwrap_or_else(|| Expression::int(1));
        stmts.push(Statement::new(StmtKind::CloseFile(Some(file_number))));
    }

    if stmts.len() == 1 {
        Ok(stmts.remove(0).kind)
    } else {
        Ok(StmtKind::Block(stmts))
    }
}

fn walk_read_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let source = pair.as_str().to_ascii_lowercase();
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut file_name = String::new();
    let mut into_var: Option<String> = None;
    let mut fail_body = Vec::new();
    let mut success_body = Vec::new();
    let mut saw_at_end = false;
    let mut saw_invalid_key = false;

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
                saw_at_end = true;
                for ac in child.into_inner() {
                    if matches!(
                        ac.as_rule(),
                        Rule::statement_list | Rule::clause_statement_list
                    ) {
                        if fail_body.is_empty() {
                            walk_statement_list(ac, &mut fail_body, ctx)?;
                        } else {
                            walk_statement_list(ac, &mut success_body, ctx)?;
                        }
                    }
                }
            }
            Rule::invalid_key_clause => {
                let mut saw_invalid_key_body = false;
                for clause in child.into_inner() {
                    if matches!(
                        clause.as_rule(),
                        Rule::statement_list | Rule::clause_statement_list
                    ) {
                        saw_invalid_key_body = true;
                        if fail_body.is_empty() {
                            walk_statement_list(clause, &mut fail_body, ctx)?;
                        } else {
                            walk_statement_list(clause, &mut success_body, ctx)?;
                        }
                    }
                }
                saw_invalid_key = saw_invalid_key_body;
            }
            _ => {}
        }
    }

    if saw_at_end {
        if let Some(status_stmt) = cobol_file_status_assign(
            ctx.file_binding(&file_name)
                .and_then(|binding| binding.status_var.as_deref()),
            "10",
        ) {
            fail_body.insert(0, status_stmt);
        }
    } else if saw_invalid_key {
        if let Some(status_stmt) = cobol_file_status_assign(
            ctx.file_binding(&file_name)
                .and_then(|binding| binding.status_var.as_deref()),
            "23",
        ) {
            fail_body.insert(0, status_stmt);
        }
    }

    let target_name = into_var
        .or_else(|| {
            ctx.file_binding(&file_name)
                .and_then(|binding| binding.record_name.clone())
        })
        .unwrap_or_else(|| file_name.clone());
    let binding = ctx
        .file_binding(&file_name)
        .cloned()
        .unwrap_or_else(|| CobolFileBinding {
            path: Expression::string(&file_name.to_ascii_lowercase()),
            file_number: 1,
            status_var: None,
            key_name: None,
            record_name: Some(target_name.clone()),
            record_fields: Vec::new() });

    let record_fields = if binding.record_fields.is_empty() {
        ctx.record_fields_for_name(&target_name)
    } else {
        binding.record_fields.clone()
    };

    let use_record_fields = !record_fields.is_empty();
    let first_record_field = record_fields.first().map(|field| field.name.clone());
    let keyed_read = use_record_fields
        && (saw_invalid_key || (!source.contains(" next") && binding.key_name.is_some()));
    let keyed_read_name = binding
        .key_name
        .as_deref()
        .or(first_record_field.as_deref())
        .unwrap_or(&target_name);
    let key_index = cobol_record_key_index(&binding, &record_fields, Some(keyed_read_name));
    let key_value = keyed_read.then(|| Expression::ident(keyed_read_name));

    if let Some(status_stmt) = cobol_file_status_assign(binding.status_var.as_deref(), "00") {
        success_body.insert(0, status_stmt);
    }

    let eof_cond = if use_record_fields {
        binary(
            BinOp::Eq,
            Expression::ident(first_record_field.as_deref().unwrap_or(&target_name)),
            Expression::null(),
        )
    } else {
        binary(
            BinOp::Eq,
            Expression::ident(&target_name),
            Expression::string(""),
        )
    };
    let else_body = (!success_body.is_empty()).then_some(success_body);

    let read_stmt = if use_record_fields {
        Statement::new(StmtKind::InputRecordFile {
            file_number: Expression::int(binding.file_number as i64),
            variables: record_fields
                .iter()
                .map(|field| field.name.clone())
                .collect(),
            key_index: keyed_read.then_some(key_index),
            key_value })
    } else {
        Statement::new(StmtKind::LineInput {
            file_number: Expression::int(binding.file_number as i64),
            variable: target_name.clone() })
    };

    Ok(StmtKind::Block(vec![
        read_stmt,
        Statement::new(StmtKind::If {
            cond: eof_cond,
            then_body: fail_body,
            elifs: Vec::new(),
            else_body }),
    ]))
}

fn walk_write_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let parts = pair.into_inner();
    let mut record_name = String::new();
    let mut source_expr: Option<Expression> = None;

    for p in parts {
        match p.as_rule() {
            Rule::ident_name => {
                if record_name.is_empty() {
                    record_name = p.as_str().to_string();
                }
            }
            Rule::write_source => {
                source_expr = Some(walk_write_source(p)?);
            }
            _ => {}
        }
    }

    let binding = ctx
        .file_binding_for_record(&record_name)
        .cloned()
        .or_else(|| ctx.file_binding(&record_name).cloned());
    let file_number = binding
        .as_ref()
        .map(|binding| Expression::int(binding.file_number as i64))
        .unwrap_or_else(|| Expression::int(1));

    let record_fields = binding
        .as_ref()
        .map(|binding| binding.record_fields.clone())
        .unwrap_or_default();
    let write_from_record = source_expr.is_none() && !record_fields.is_empty();

    let write_stmt = if write_from_record {
        Statement::new(StmtKind::WriteFile {
            file_number,
            items: record_fields
                .iter()
                .map(|field| Expression::ident(&field.name))
                .collect() })
    } else {
        let data_expr = source_expr.unwrap_or_else(|| Expression::ident(&record_name));
        let group_items = match &data_expr.kind {
            ExprKind::Ident(name) => ctx.group_layout_for_name(name),
            _ => Vec::new() };
        let flattened_expr = if group_items.is_empty() {
            data_expr
        } else {
            let mut items = group_items.into_iter();
            let mut result = items.next().unwrap_or_else(|| Expression::string(""));
            for item in items {
                result = binary(BinOp::Concat, result, item);
            }
            result
        };
        Statement::new(StmtKind::PrintFile {
            file_number,
            items: vec![flattened_expr] })
    };

    let mut stmts = vec![write_stmt];
    if let Some(status_stmt) =
        cobol_file_status_assign(binding.as_ref().and_then(|b| b.status_var.as_deref()), "00")
    {
        stmts.push(status_stmt);
    }
    Ok(if stmts.len() == 1 {
        stmts.remove(0).kind
    } else {
        StmtKind::Block(stmts)
    })
}

fn walk_rewrite_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let fallback_pair = pair.clone();
    let parts = pair.into_inner();
    let mut record_name = String::new();
    let mut source_expr: Option<Expression> = None;

    for p in parts {
        match p.as_rule() {
            Rule::ident_name => {
                if record_name.is_empty() {
                    record_name = p.as_str().to_string();
                }
            }
            Rule::write_source => {
                source_expr = Some(walk_write_source(p)?);
            }
            _ => {}
        }
    }

    let binding = ctx
        .file_binding_for_record(&record_name)
        .cloned()
        .or_else(|| ctx.file_binding(&record_name).cloned());
    let file_number = binding
        .as_ref()
        .map(|binding| Expression::int(binding.file_number as i64))
        .unwrap_or_else(|| Expression::int(1));
    let record_fields = binding
        .as_ref()
        .map(|binding| binding.record_fields.clone())
        .unwrap_or_default();

    let rewrite_stmt = if source_expr.is_none() && !record_fields.is_empty() {
        Statement::new(StmtKind::RewriteRecordFile {
            file_number,
            items: record_fields
                .iter()
                .map(|field| Expression::ident(&field.name))
                .collect(),
            field_formats: record_fields
                .iter()
                .map(|field| {
                    field.numeric.then_some(RecordFieldFormat {
                        decimal_places: field.decimal_places })
                })
                .collect() })
    } else {
        return walk_write_stmt(fallback_pair, ctx);
    };

    let mut stmts = vec![rewrite_stmt];
    if let Some(status_stmt) =
        cobol_file_status_assign(binding.as_ref().and_then(|b| b.status_var.as_deref()), "00")
    {
        stmts.push(status_stmt);
    }
    Ok(if stmts.len() == 1 {
        stmts.remove(0).kind
    } else {
        StmtKind::Block(stmts)
    })
}

fn walk_start_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let source = pair.as_str().to_string();
    let mut file_name = String::new();
    let mut key_name: Option<String> = None;

    for child in pair.into_inner() {
        if child.as_rule() == Rule::ident_name {
            if file_name.is_empty() {
                file_name = child.as_str().to_string();
            } else {
                key_name = Some(child.as_str().to_string());
            }
        }
    }

    let Some(binding) = ctx.file_binding(&file_name).cloned() else {
        return Ok(StmtKind::Empty);
    };
    let record_name = binding
        .record_name
        .clone()
        .unwrap_or_else(|| file_name.clone());
    let record_fields = if binding.record_fields.is_empty() {
        ctx.record_fields_for_name(&record_name)
    } else {
        binding.record_fields.clone()
    };
    let relation = parse_cobol_start_relation(&source);
    let Some(key_name) = key_name.or_else(|| binding.key_name.clone()) else {
        return Ok(StmtKind::Empty);
    };
    let key_index = cobol_record_key_index(&binding, &record_fields, Some(&key_name));

    let mut stmts = vec![Statement::new(StmtKind::StartFile {
        file_number: Expression::int(binding.file_number as i64),
        key_index,
        key_value: Expression::ident(&key_name),
        relation })];
    if let Some(status_stmt) = cobol_file_status_assign(binding.status_var.as_deref(), "00") {
        stmts.push(status_stmt);
    }
    Ok(if stmts.len() == 1 {
        stmts.remove(0).kind
    } else {
        StmtKind::Block(stmts)
    })
}

fn walk_write_source(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut parts = Vec::new();

    for child in pair.into_inner() {
        if child.as_rule() != Rule::write_source_part {
            continue;
        }

        let mut saw_all = false;
        let mut expr: Option<Expression> = None;
        for part in child.into_inner() {
            match part.as_rule() {
                Rule::kw_all => {
                    saw_all = true;
                }
                Rule::expression => {
                    expr = Some(walk_expression(part)?);
                }
                Rule::literal => {
                    expr = Some(walk_literal(part)?);
                }
                _ => {}
            }
        }

        let mut part_expr = expr.unwrap_or_else(|| Expression::string(""));
        if saw_all {
            // Best-effort support for `FROM ALL literal`; preserve the fill
            // character even though fixed-width record padding happens later.
            part_expr = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(part_expr),
                    field: "repeat".to_string(),
                    null_safe: false })),
                args: vec![Argument::positional(Expression::int(80))],
                optional: false });
        }
        parts.push(part_expr);
    }

    if parts.is_empty() {
        return Ok(Expression::string(""));
    }

    let mut result = parts.remove(0);
    for part in parts {
        result = binary(BinOp::Concat, result, part);
    }
    Ok(result)
}

fn walk_delete_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
    let parts = inner_nokw(pair);
    let name = parts
        .into_iter()
        .find(|p| p.as_rule() == Rule::ident_name)
        .map(|p| p.as_str().to_string())
        .unwrap_or_default();
    if let Some(status_stmt) = cobol_file_status_assign(
        ctx.file_binding(&name)
            .and_then(|binding| binding.status_var.as_deref()),
        "00",
    ) {
        Ok(StmtKind::Block(vec![status_stmt]))
    } else {
        Ok(StmtKind::Empty)
    }
}

// ── SEARCH ──────────────────────────────────────────────────────────────────

fn walk_search_stmt(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<StmtKind, String> {
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
                let cond = walk_condition(child, ctx)?;
                when_clauses.push((cond, Vec::new()));
            }
            Rule::statement_list => {
                if let Some(last) = when_clauses.last_mut() {
                    walk_statement_list(child, &mut last.1, ctx)?;
                } else {
                    walk_statement_list(child, &mut at_end_body, ctx)?;
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
    let else_body = if at_end_body.is_empty() {
        None
    } else {
        Some(at_end_body)
    };

    Ok(StmtKind::If {
        cond: first_cond,
        then_body: first_body,
        elifs,
        else_body })
}

// ── INVOKE (OO COBOL) ──────────────────────────────────────────────────────

fn walk_invoke_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut obj = String::new();
    let mut method = String::new();
    let mut args: Vec<Argument> = Vec::new();
    let mut returning_var: Option<String> = None;
    let mut in_using = false;
    let mut in_returning = false;

    for child in children {
        match child.as_rule() {
            Rule::kw_using => {
                in_using = true;
                in_returning = false;
            }
            Rule::kw_returning => {
                in_returning = true;
                in_using = false;
            }
            Rule::kw_new => {
                method = "new".to_string();
            }
            Rule::ident_name => {
                let name = child.as_str().to_string();
                if obj.is_empty() {
                    obj = name;
                } else if in_returning {
                    returning_var = Some(name);
                } else if method.is_empty() {
                    method = name;
                } else if in_using {
                    args.push(Argument::positional(Expression::ident(&name)));
                }
            }
            _ => {}
        }
    }

    let call_expr = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&obj)),
            field: method,
            null_safe: false })),
        args,
        optional: false });

    if let Some(ret_var) = returning_var {
        Ok(StmtKind::Assign {
            targets: vec![Expression::ident(&ret_var)],
            value: call_expr, by_ref: false })
    } else {
        Ok(StmtKind::Expr(call_expr))
    }
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
                    name = s[1..s.len() - 1].to_string();
                }
            }
            Rule::ident_name => {
                if name.is_empty() {
                    name = child.as_str().to_string();
                } else if in_using {
                    args.push(Argument::positional(Expression::ident(child.as_str())));
                }
            }
            Rule::kw_using => {
                in_using = true;
            }
            _ => {}
        }
    }

    Ok(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(&name)),
        args,
        optional: false })))
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
            decorators: vec![] },
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
                let mut local_ctx = CobolWalkerContext::new();
                walk_data_division(child, &mut field_stmts, &mut local_ctx)?;
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
                                    array_bounds: None });
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
                        if let StmtKind::FunctionDecl {
                            ref mut modifiers, ..
                        } = method.kind
                        {
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
                            let mut local_ctx = CobolWalkerContext::new();
                            walk_data_division(oc, &mut field_stmts, &mut local_ctx)?;
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
                                                array_bounds: None });
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
    let mut method_ctx = CobolWalkerContext::new();

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
                        Rule::kw_override => {
                            is_override = true;
                        }
                        Rule::property_modifier => {
                            for pm in m.into_inner() {
                                match pm.as_rule() {
                                    Rule::kw_get => {
                                        is_property_get = true;
                                    }
                                    Rule::kw_set => {
                                        is_property_set = true;
                                    }
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
                        walk_data_item(md, &mut body, &mut method_ctx)?;
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
                            walk_statement_list(mp, &mut body, &method_ctx)?;
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
    let is_generator = body_has_yield(&body);

    Ok(Statement::with_span(
        StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            body,
            modifiers,
            handles: Vec::new(),
            is_async: false,
            is_generator,
            is_sub },
        span,
    ))
}

fn walk_using_params(pair: Pair<Rule>) -> Vec<Param> {
    let mut params = Vec::new();
    let mut pass_by = PassBy::Value;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::kw_reference => {
                pass_by = PassBy::Ref;
            }
            Rule::kw_content => {
                pass_by = PassBy::Const;
            }
            Rule::ident_name => {
                params.push(Param {
                    name: child.as_str().to_string(),
                    type_hint: None,
                    default: None,
                    pass_by,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false });
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
            decorators: vec![] },
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
        signature_source: None })
}

// ════════════════════════════════════════════════════════════════════════════
// Conditions
// ════════════════════════════════════════════════════════════════════════════

fn walk_condition(pair: Pair<Rule>, ctx: &CobolWalkerContext) -> Result<Expression, String> {
    let inner = pair.into_inner().next().ok_or("empty condition")?;
    Ok(rewrite_condition_names(walk_or_condition(inner)?, ctx))
}

/// COBOL compares a numeric DISPLAY item against an alphanumeric literal by
/// its CHARACTER representation, not its value: a `PIC 9(3)` holding 6 equals
/// `"006"`. Without this the value 6 is compared against the string "006" and
/// the condition is false, where `cobc` says true — measured, and it silently
/// inverted every such test.
///
/// Only NUMERIC pictures are realigned. An alphanumeric field already compares
/// correctly (`PIC X(10)` holding "hi" equals `"hi"`, because COBOL pads the
/// shorter operand), and padding it here would break that.
fn cobol_align_pic_comparison(
    op: BinOp,
    left: Expression,
    right: Expression,
    ctx: &CobolWalkerContext,
) -> (Expression, Expression) {
    if !matches!(op, BinOp::Eq | BinOp::NotEq) {
        return (left, right);
    }
    let numeric_pic = |e: &Expression| -> Option<CobolPicFmt> {
        let ExprKind::Ident(name) = &e.kind else {
            return None;
        };
        match ctx.field_pic(name) {
            Some(fmt @ CobolPicFmt::Numeric(_)) => Some(fmt),
            _ => None }
    };
    let is_text = |e: &Expression| matches!(&e.kind, ExprKind::Lit(Literal::Str(_)));

    if is_text(&right) && let Some(fmt) = numeric_pic(&left) {
        return (cobol_pic_format_expr(left, fmt), right);
    }
    if is_text(&left) && let Some(fmt) = numeric_pic(&right) {
        return (left, cobol_pic_format_expr(right, fmt));
    }
    (left, right)
}

fn rewrite_condition_names(expr: Expression, ctx: &CobolWalkerContext) -> Expression {
    let span = expr.span;
    let kind = match expr.kind {
        ExprKind::Ident(name) => {
            return ctx
                .condition_expr(&name)
                .map(|resolved| Expression::with_span(resolved.kind, span))
                .unwrap_or_else(|| Expression::with_span(ExprKind::Ident(name), span));
        }
        ExprKind::Binary { op, left, right } => {
            let left = rewrite_condition_names(*left, ctx);
            let right = rewrite_condition_names(*right, ctx);
            let (left, right) = cobol_align_pic_comparison(op, left, right, ctx);
            ExprKind::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right) }
        }
        ExprKind::Unary { op, expr } => ExprKind::Unary {
            op,
            expr: Box::new(rewrite_condition_names(*expr, ctx)) },
        ExprKind::Call {
            callee,
            args,
            optional } => ExprKind::Call {
            callee: Box::new(rewrite_condition_names(*callee, ctx)),
            args: args
                .into_iter()
                .map(|arg| Argument {
                    value: rewrite_condition_names(arg.value, ctx),
                    ..arg
                })
                .collect(),
            optional },
        ExprKind::Member {
            object,
            field,
            null_safe } => ExprKind::Member {
            object: Box::new(rewrite_condition_names(*object, ctx)),
            field,
            null_safe },
        ExprKind::Index {
            object,
            index,
            null_safe } => ExprKind::Index {
            object: Box::new(rewrite_condition_names(*object, ctx)),
            index: Box::new(rewrite_condition_names(*index, ctx)),
            null_safe },
        ExprKind::Ternary { cond, then, else_ } => ExprKind::Ternary {
            cond: Box::new(rewrite_condition_names(*cond, ctx)),
            then: Box::new(rewrite_condition_names(*then, ctx)),
            else_: Box::new(rewrite_condition_names(*else_, ctx)) },
        ExprKind::NullCoalesce { left, right } => ExprKind::NullCoalesce {
            left: Box::new(rewrite_condition_names(*left, ctx)),
            right: Box::new(rewrite_condition_names(*right, ctx)) },
        other => other };
    Expression::with_span(kind, span)
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
    let comparison = children
        .iter()
        .find(|c| c.as_rule() == Rule::comparison)
        .ok_or("not_condition without comparison")?;

    let expr = walk_comparison(comparison.clone())?;

    if has_not {
        Ok(Expression::new(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(expr) }))
    } else {
        Ok(expr)
    }
}

fn walk_comparison(pair: Pair<Rule>) -> Result<Expression, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Check for class condition test (IS NUMERIC, etc.)
    if let Some(class_test) = children
        .iter()
        .find(|c| c.as_rule() == Rule::class_condition_test)
    {
        let expr = walk_expression(children.first().unwrap().clone())?;
        return walk_class_condition_test(class_test.clone(), expr);
    }

    // Check for sign condition test (IS POSITIVE, etc.)
    if let Some(sign_test) = children
        .iter()
        .find(|c| c.as_rule() == Rule::sign_condition_test)
    {
        let expr = walk_expression(children.first().unwrap().clone())?;
        return walk_sign_condition_test(sign_test.clone(), expr);
    }

    // Check for comparison operator
    if let Some(comp_op) = children.iter().find(|c| c.as_rule() == Rule::comparison_op) {
        let exprs: Vec<Expression> = children
            .iter()
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

    let call = if children.iter().any(|c| c.as_rule() == Rule::kw_numeric) {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__is_numeric")),
            args: vec![Argument::positional(expr)],
            optional: false })
    } else if children
        .iter()
        .any(|c| c.as_rule() == Rule::kw_alphabetic_lower)
    {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__is_alphabetic_lower")),
            args: vec![Argument::positional(expr)],
            optional: false })
    } else if children
        .iter()
        .any(|c| c.as_rule() == Rule::kw_alphabetic_upper)
    {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__is_alphabetic_upper")),
            args: vec![Argument::positional(expr)],
            optional: false })
    } else if let Some(class_name) = children.iter().find(|c| c.as_rule() == Rule::ident_name) {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__is_class")),
            args: vec![
                Argument::positional(expr),
                Argument::positional(Expression::string(class_name.as_str())),
            ],
            optional: false })
    } else {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__is_alphabetic")),
            args: vec![Argument::positional(expr)],
            optional: false })
    };

    if is_negated {
        Ok(negate_expr(call))
    } else {
        Ok(call)
    }
}

fn walk_sign_condition_test(pair: Pair<Rule>, expr: Expression) -> Result<Expression, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    let is_negated = children.iter().any(|c| c.as_rule() == Rule::kw_not);

    let result = if children
        .iter()
        .any(|c| matches!(c.as_rule(), Rule::kw_positive))
    {
        binary(BinOp::Gt, expr, Expression::int(0))
    } else if children
        .iter()
        .any(|c| matches!(c.as_rule(), Rule::kw_negative))
    {
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
    let raw_text = pair.as_str().trim().to_uppercase();
    let children: Vec<Pair<Rule>> = pair.clone().into_inner().collect();
    let text = if children.is_empty() {
        raw_text.replace("\t", " ")
    } else {
        children
            .iter()
            .map(|c| c.as_str().to_uppercase())
            .collect::<Vec<_>>()
            .join(" ")
    };
    let is_negated = raw_text.contains(" NOT ")
        || raw_text.starts_with("NOT ")
        || children.iter().any(|c| c.as_rule() == Rule::kw_not);

    // Check for symbolic operators first. Pest does not emit literal operator
    // tokens like `<` and `<=` as child pairs, so use the outer span text.
    if raw_text.contains(">=") {
        return BinOp::GtEq;
    }
    if raw_text.contains("<=") {
        return BinOp::LtEq;
    }
    if raw_text.contains(">") {
        return if is_negated { BinOp::LtEq } else { BinOp::Gt };
    }
    if raw_text.contains("<") {
        return if is_negated { BinOp::GtEq } else { BinOp::Lt };
    }
    if raw_text.contains("=") && is_negated {
        return BinOp::NotEq;
    }
    if raw_text.contains("=") {
        return BinOp::Eq;
    }

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
            let op = if op_text == "+" {
                BinOp::Add
            } else {
                BinOp::Sub
            };
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
            let op = if op_text.starts_with('*') {
                BinOp::Mul
            } else {
                BinOp::Div
            };
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
                expr: Box::new(expr) }));
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

        Rule::function_call => walk_function_call(pair),

        Rule::new_expr => walk_new_expr(pair),

        Rule::paren_expr => {
            let inner = pair.into_inner().next().ok_or("empty paren_expr")?;
            walk_expression(inner)
        }

        Rule::literal => walk_literal(pair),

        Rule::qualified_ident => walk_qualified_ident(pair),

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
                Ok(Expression::with_span(
                    ExprKind::Ident(text.to_string()),
                    span,
                ))
            } else {
                Err(format!("COBOL walker: unhandled atom rule {:?}", other))
            }
        }
    }
}

// ── Function calls ──────────────────────────────────────────────────────────

/// Lower COBOL numeric intrinsics that have no direct host mapping into
/// expression trees composed from primitives the profile already resolves
/// (arithmetic ops plus MAX/MIN/POWER/SQRT/f64_floor/f64_trunc). Returns None
/// for names handled directly by the profile.
fn desugar_cobol_math_intrinsic(name: &str, args: &[Argument]) -> Option<Expression> {
    let xs: Vec<Expression> = args.iter().map(|a| a.value.clone()).collect();
    let n = xs.len();

    let call = |fname: &str, cargs: Vec<Expression>| -> Expression {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(fname)),
            args: cargs.into_iter().map(Argument::positional).collect(),
            optional: false })
    };
    let sum = |vals: &[Expression]| -> Expression {
        let mut it = vals.iter().cloned();
        let first = it.next().expect("sum needs >=1 arg");
        it.fold(first, |acc, x| binary(BinOp::Add, acc, x))
    };
    // Format a value COBOL-style for financial intrinsics: truncate to two
    // decimal places and render with a trailing-zero 2-decimal string.
    let fmt2 = |v: Expression| -> Expression {
        let ncall = |fname: &str, arg: Expression| -> Expression {
            Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(fname)),
                args: vec![Argument::positional(arg)],
                optional: false })
        };
        // Binary floats can't represent e.g. 1.10 exactly, so raw truncation of
        // 1.0999999 yields 1.09. Round to 6 dp first to erase that noise, then
        // truncate to 2 dp (COBOL fixed-point behaviour).
        let denoise = binary(
            BinOp::Div,
            ncall(
                "f64_nearest",
                binary(BinOp::Mul, v, Expression::float(1_000_000.0)),
            ),
            Expression::float(1_000_000.0),
        );
        let scaled = ncall(
            "f64_trunc",
            binary(BinOp::Mul, denoise, Expression::int(100)),
        );
        let back = binary(BinOp::Div, scaled, Expression::int(100));
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__to_fixed2")),
            args: vec![
                Argument::positional(back),
                Argument::positional(Expression::int(2)),
            ],
            optional: false })
    };
    // Population variance: mean of squared deviations from the mean.
    let pop_variance = |vals: &[Expression]| -> Expression {
        let mean = binary(BinOp::Div, sum(vals), Expression::int(vals.len() as i64));
        let sq_devs: Vec<Expression> = vals
            .iter()
            .map(|x| {
                let d = binary(BinOp::Sub, x.clone(), mean.clone());
                binary(BinOp::Mul, d.clone(), d)
            })
            .collect();
        binary(
            BinOp::Div,
            sum(&sq_devs),
            Expression::int(vals.len() as i64),
        )
    };

    match name {
        "PI" if n == 0 => Some(Expression::float(std::f64::consts::PI)),
        "E" if n == 0 => Some(Expression::float(std::f64::consts::E)),
        "SUM" if n >= 1 => Some(sum(&xs)),
        "MEAN" | "AVERAGE" if n >= 1 => {
            Some(binary(BinOp::Div, sum(&xs), Expression::int(n as i64)))
        }
        "RANGE" if n >= 2 => Some(binary(
            BinOp::Sub,
            call("MAX", xs.clone()),
            call("MIN", xs.clone()),
        )),
        "MIDRANGE" if n >= 2 => Some(binary(
            BinOp::Div,
            binary(BinOp::Add, call("MAX", xs.clone()), call("MIN", xs.clone())),
            Expression::int(2),
        )),
        // MOD is floor-based; REM truncates toward zero.
        "MOD" if n == 2 => {
            let (a, b) = (xs[0].clone(), xs[1].clone());
            let q = call("f64_floor", vec![binary(BinOp::Div, a.clone(), b.clone())]);
            Some(binary(BinOp::Sub, a, binary(BinOp::Mul, b, q)))
        }
        "REM" if n == 2 => {
            let (a, b) = (xs[0].clone(), xs[1].clone());
            let q = call("f64_trunc", vec![binary(BinOp::Div, a.clone(), b.clone())]);
            Some(binary(BinOp::Sub, a, binary(BinOp::Mul, b, q)))
        }
        // COBOL ORD is 1-based (ORD('A') == 66); CHAR is its inverse.
        "ORD" if n == 1 => Some(binary(
            BinOp::Add,
            call("__ord_code", vec![xs[0].clone()]),
            Expression::int(1),
        )),
        "CHAR" if n == 1 => Some(call(
            "__char_code",
            vec![binary(BinOp::Sub, xs[0].clone(), Expression::int(1))],
        )),
        "INTEGER-PART" if n == 1 => Some(call("f64_trunc", vec![xs[0].clone()])),
        "EXP10" if n == 1 => Some(call("POWER", vec![Expression::int(10), xs[0].clone()])),
        "VARIANCE" if n >= 1 => Some(pop_variance(&xs)),
        "STANDARD-DEVIATION" if n >= 1 => Some(call("SQRT", vec![pop_variance(&xs)])),
        // ANNUITY(rate, periods): rate==0 ? 1/periods : rate/(1-(1+rate)^-periods)
        "ANNUITY" if n == 2 => {
            let (rate, periods) = (xs[0].clone(), xs[1].clone());
            let one_plus = binary(BinOp::Add, Expression::int(1), rate.clone());
            let neg_n = binary(BinOp::Sub, Expression::int(0), periods.clone());
            let denom = binary(
                BinOp::Sub,
                Expression::int(1),
                call("POWER", vec![one_plus, neg_n]),
            );
            let general = binary(BinOp::Div, rate.clone(), denom);
            let zero_case = binary(BinOp::Div, Expression::int(1), periods);
            Some(fmt2(Expression::new(ExprKind::Ternary {
                cond: Box::new(binary(BinOp::Eq, rate, Expression::int(0))),
                then: Box::new(zero_case),
                else_: Box::new(general) })))
        }
        // PRESENT-VALUE(rate, v1..vk): sum of vi / (1+rate)^i for i = 1..k
        "PRESENT-VALUE" if n >= 2 => {
            let rate = xs[0].clone();
            let one_plus = binary(BinOp::Add, Expression::int(1), rate);
            let terms: Vec<Expression> = xs[1..]
                .iter()
                .enumerate()
                .map(|(i, v)| {
                    let disc = call(
                        "POWER",
                        vec![one_plus.clone(), Expression::int(i as i64 + 1)],
                    );
                    binary(BinOp::Div, v.clone(), disc)
                })
                .collect();
            Some(fmt2(sum(&terms)))
        }
        _ => None }
}

/// COBOL integer-date intrinsics. Day 1 = 1601-01-01; the Unix epoch sits at
/// integer 134775 (days since 1601-01-01, 1-based). Conversions compose
/// `ecma:date` UTC getters + `Date.UTC` with plain arithmetic — no host date
/// functions specific to COBOL.
fn desugar_cobol_date_intrinsic(name: &str, args: &[Argument]) -> Option<Expression> {
    fn dcall(f: &str, a: Vec<Expression>) -> Expression {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(f)),
            args: a.into_iter().map(Argument::positional).collect(),
            optional: false })
    }
    fn didiv(a: Expression, b: i64) -> Expression {
        dcall("f64_trunc", vec![binary(BinOp::Div, a, Expression::int(b))])
    }
    fn imod(a: Expression, b: i64) -> Expression {
        binary(
            BinOp::Sub,
            a.clone(),
            binary(BinOp::Mul, Expression::int(b), didiv(a, b)),
        )
    }
    // Unix-epoch milliseconds for a 1-based COBOL integer date.
    fn ms_from_int(e: Expression) -> Expression {
        binary(
            BinOp::Mul,
            binary(BinOp::Sub, e, Expression::int(134_775)),
            Expression::int(86_400_000),
        )
    }
    fn tern(c: Expression, t: Expression, e: Expression) -> Expression {
        Expression::new(ExprKind::Ternary {
            cond: Box::new(c),
            then: Box::new(t),
            else_: Box::new(e) })
    }
    // `lo <= v <= hi` violated → the out-of-range branch.
    fn out_of_range(v: Expression, lo: i64, hi: i64) -> Expression {
        binary(
            BinOp::Or,
            binary(BinOp::Lt, v.clone(), Expression::int(lo)),
            binary(BinOp::Gt, v, Expression::int(hi)),
        )
    }
    // Integer-of-date of Jan 1 of `year`: UTC(year,0,1)/86_400_000 + 134775.
    fn jan1_int(year: Expression) -> Expression {
        binary(
            BinOp::Add,
            didiv(
                dcall(
                    "__date_utc",
                    vec![year, Expression::int(0), Expression::int(1)],
                ),
                86_400_000,
            ),
            Expression::int(134_775),
        )
    }

    let xs: Vec<Expression> = args.iter().map(|a| a.value.clone()).collect();
    if xs.len() != 1 {
        return None;
    }
    let arg = xs[0].clone();
    match name {
        "DATE-OF-INTEGER" => {
            let ms = ms_from_int(arg);
            let y = dcall("__utc_year", vec![ms.clone()]);
            let m = binary(
                BinOp::Add,
                dcall("__utc_month", vec![ms.clone()]),
                Expression::int(1),
            );
            let d = dcall("__utc_date", vec![ms]);
            Some(binary(
                BinOp::Add,
                binary(
                    BinOp::Add,
                    binary(BinOp::Mul, y, Expression::int(10_000)),
                    binary(BinOp::Mul, m, Expression::int(100)),
                ),
                d,
            ))
        }
        "INTEGER-OF-DAY" => {
            // year = yyyyddd / 1000; ddd = yyyyddd % 1000 → jan1_int(year)+ddd-1
            let year = didiv(arg.clone(), 1000);
            let ddd = imod(arg, 1000);
            Some(binary(
                BinOp::Sub,
                binary(BinOp::Add, jan1_int(year), ddd),
                Expression::int(1),
            ))
        }
        "DAY-OF-INTEGER" => {
            let year = dcall("__utc_year", vec![ms_from_int(arg.clone())]);
            // ddd = n - jan1_int(year) + 1; result = year*1000 + ddd
            let ddd = binary(
                BinOp::Add,
                binary(BinOp::Sub, arg, jan1_int(year.clone())),
                Expression::int(1),
            );
            Some(binary(
                BinOp::Add,
                binary(BinOp::Mul, year, Expression::int(1000)),
                ddd,
            ))
        }
        // Validation: 0 = valid, 1 = year out of [1601,9999], 2 = bad month,
        // 3 = bad day (ecma rolls an invalid day into the next month, so the
        // round-tripped day differs).
        "TEST-DATE-YYYYMMDD" => {
            let y = didiv(arg.clone(), 10_000);
            let m = didiv(imod(arg.clone(), 10_000), 100);
            let day = imod(arg, 100);
            let ms = dcall(
                "__date_utc",
                vec![
                    y.clone(),
                    binary(BinOp::Sub, m.clone(), Expression::int(1)),
                    day.clone(),
                ],
            );
            let bad_day = binary(BinOp::NotEq, dcall("__utc_date", vec![ms]), day);
            Some(tern(
                out_of_range(y, 1601, 9999),
                Expression::int(1),
                tern(
                    out_of_range(m, 1, 12),
                    Expression::int(2),
                    tern(bad_day, Expression::int(3), Expression::int(0)),
                ),
            ))
        }
        // 0 = valid, 1 = year out of range, 2 = day-of-year out of [1, days-in-year].
        "TEST-DAY-YYYYDDD" => {
            let y = didiv(arg.clone(), 1000);
            let ddd = imod(arg, 1000);
            // days-in-year(y) = (UTC(y+1,0,1) - UTC(y,0,1)) / 86_400_000
            let diy = didiv(
                binary(
                    BinOp::Sub,
                    dcall(
                        "__date_utc",
                        vec![
                            binary(BinOp::Add, y.clone(), Expression::int(1)),
                            Expression::int(0),
                            Expression::int(1),
                        ],
                    ),
                    dcall(
                        "__date_utc",
                        vec![y.clone(), Expression::int(0), Expression::int(1)],
                    ),
                ),
                86_400_000,
            );
            let bad_ddd = binary(
                BinOp::Or,
                binary(BinOp::Lt, ddd.clone(), Expression::int(1)),
                binary(BinOp::Gt, ddd, diy),
            );
            Some(tern(
                out_of_range(y, 1601, 9999),
                Expression::int(1),
                tern(bad_ddd, Expression::int(2), Expression::int(0)),
            ))
        }
        _ => None }
}

/// Reduce a COBOL intrinsic over a whole OCCURS table (`FUNCTION MAX(ALL VALS)`).
/// The table is an `ObjectKind::Array`, identical to a JS/Python sequence, so the
/// aggregates route to the shared ECMA iterable reducers
/// (`__table_max/min/sum` → `ecma:math:maxOf/minOf/sumPrecise`) — the exact
/// surface Python's `max`/`min`/`sum` use. Composite stats build from those.
fn desugar_cobol_table_aggregate(name: &str, arr: Expression) -> Option<Expression> {
    let call1 = |fname: &str, a: Expression| -> Expression {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(fname)),
            args: vec![Argument::positional(a)],
            optional: false })
    };
    let length = |a: Expression| -> Expression {
        Expression::new(ExprKind::Member {
            object: Box::new(a),
            field: "length".to_string(),
            null_safe: false })
    };

    match name {
        "MAX" => Some(call1("__table_max", arr)),
        "MIN" => Some(call1("__table_min", arr)),
        "SUM" => Some(call1("__table_sum", arr)),
        "MEAN" | "AVERAGE" => Some(binary(
            BinOp::Div,
            call1("__table_sum", arr.clone()),
            length(arr),
        )),
        "RANGE" => Some(binary(
            BinOp::Sub,
            call1("__table_max", arr.clone()),
            call1("__table_min", arr),
        )),
        "MIDRANGE" => Some(binary(
            BinOp::Div,
            binary(
                BinOp::Add,
                call1("__table_max", arr.clone()),
                call1("__table_min", arr),
            ),
            Expression::int(2),
        )),
        _ => None }
}

fn walk_function_call(pair: Pair<Rule>) -> Result<Expression, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();

    let mut func_name = String::new();
    let mut args: Vec<Argument> = Vec::new();
    let mut subscript_or_refmod: Option<Pair<Rule>> = None;
    // Set when an argument is `ALL table-name` — the whole OCCURS array.
    let mut all_table: Option<Expression> = None;

    for child in children {
        match child.as_rule() {
            Rule::function_name => {
                func_name = child.as_str().to_uppercase();
            }
            Rule::function_call_args => {
                for arg_child in child.into_inner() {
                    match arg_child.as_rule() {
                        Rule::func_args => {
                            for func_arg in arg_child.into_inner() {
                                match func_arg.as_rule() {
                                    Rule::all_table_ref => {
                                        for inner in func_arg.into_inner() {
                                            if inner.as_rule() == Rule::qualified_ident {
                                                all_table = Some(walk_qualified_ident(inner)?);
                                            }
                                        }
                                    }
                                    Rule::func_kw_arg => {
                                        // Reserved-word argument (e.g. TRIM's
                                        // LEADING/TRAILING) — pass as an ident.
                                        args.push(Argument::positional(Expression::ident(
                                            &func_arg.as_str().to_ascii_uppercase(),
                                        )));
                                    }
                                    Rule::expression => {
                                        args.push(Argument::positional(walk_expression(func_arg)?));
                                    }
                                    Rule::atom => {
                                        args.push(Argument::positional(walk_atom(func_arg)?));
                                    }
                                    _ => {}
                                }
                            }
                        }
                        Rule::expression => {
                            args.push(Argument::positional(walk_expression(arg_child)?));
                        }
                        Rule::atom => {
                            args.push(Argument::positional(walk_atom(arg_child)?));
                        }
                        _ => {}
                    }
                }
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
            Rule::subscript_or_refmod => {
                subscript_or_refmod = Some(child);
            }
            _ => {}
        }
    }

    // `FUNCTION MAX(ALL VALS)` etc.: the argument is one whole OCCURS array, so
    // reduce over it via the shared ECMA iterable reducers (same surface Python
    // uses) rather than the scalar variadic form.
    if let Some(table) = all_table {
        if let Some(agg) = desugar_cobol_table_aggregate(&func_name, table.clone()) {
            return Ok(agg);
        }
        // Unhandled aggregate (e.g. MEDIAN): fall through with the array as the
        // sole argument.
        args.push(Argument::positional(table));
    }

    // TRIM(x LEADING|TRAILING) → trimStart/trimEnd; bare TRIM(x) → str_trim.
    if func_name == "TRIM" && args.len() == 2 {
        if let ExprKind::Ident(dir) = &args[1].value.kind {
            let helper = match dir.to_ascii_uppercase().as_str() {
                "LEADING" => Some("__trim_start"),
                "TRAILING" => Some("__trim_end"),
                _ => None };
            if let Some(h) = helper {
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(h)),
                    args: vec![Argument::positional(args[0].value.clone())],
                    optional: false }));
            }
        }
    }

    if subscript_or_refmod.is_none() {
        if let Some(desugared) = desugar_cobol_date_intrinsic(&func_name, &args) {
            return Ok(desugared);
        }
        if let Some(desugared) = desugar_cobol_math_intrinsic(&func_name, &args) {
            return Ok(desugared);
        }
    }

    let mut expr = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(&func_name)),
        args,
        optional: false });

    if let Some(sub_pair) = subscript_or_refmod {
        expr = apply_cobol_subscript_or_refmod(expr, sub_pair)?;
    }

    Ok(expr)
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
        args }))
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
        expr = apply_cobol_subscript_or_refmod(expr, sub_pair)?;
    }

    // Handle qualification: field OF group → group.field
    if let Some(parent) = qualification {
        expr = Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&parent)),
            field: name,
            null_safe: false });
    }

    Ok(expr)
}

fn walk_data_target_expr(pair: Pair<Rule>) -> Result<Expression, String> {
    match pair.as_rule() {
        Rule::data_target => {
            let children: Vec<Pair<Rule>> = pair.into_inner().collect();

            let mut name = String::new();
            let mut subscript: Option<Pair<Rule>> = None;
            let mut qualification: Option<String> = None;

            for child in &children {
                match child.as_rule() {
                    Rule::ident_name | Rule::kw_sd => {
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

            if let Some(sub_pair) = subscript {
                expr = apply_cobol_subscript_or_refmod(expr, sub_pair)?;
            }

            if let Some(parent) = qualification {
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(&parent)),
                    field: name,
                    null_safe: false });
            }

            Ok(expr)
        }
        Rule::ident_name | Rule::kw_sd => Ok(Expression::ident(pair.as_str())),
        Rule::qualified_ident => walk_qualified_ident(pair),
        _ => Err(format!(
            "COBOL walker: expected data_target, got {:?}",
            pair.as_rule()
        )) }
}

fn apply_cobol_subscript_or_refmod(
    mut expr: Expression,
    sub_pair: Pair<Rule>,
) -> Result<Expression, String> {
    // The grammar tags each alternative: `refmod` = name(start:length),
    // `subscript` = name(idx[, idx]...). The `:`/`,` separators are literals
    // consumed by the grammar, so the alternative is only recoverable from the
    // rule tag — never from scanning the child expressions' text.
    let inner = sub_pair
        .into_inner()
        .next()
        .ok_or("empty subscript_or_refmod")?;
    let is_refmod = inner.as_rule() == Rule::refmod;
    let expr_children: Vec<Pair<Rule>> = inner
        .into_inner()
        .filter(|c| c.as_rule() == Rule::expression)
        .collect();

    if is_refmod {
        let start_expr = if !expr_children.is_empty() {
            walk_expression(expr_children[0].clone())?
        } else {
            Expression::int(1)
        };
        let length_expr = if expr_children.len() >= 2 {
            walk_expression(expr_children[1].clone())?
        } else {
            Expression::int(0)
        };

        let adjusted_start = normalize_array_index_operand(start_expr, COBOL_ARRAY_INDEXING);

        return Ok(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__refmod")),
            args: vec![
                Argument::positional(expr),
                Argument::positional(adjusted_start),
                Argument::positional(length_expr),
            ],
            optional: false }));
    }

    if !expr_children.is_empty() {
        let index_expr = walk_expression(expr_children[0].clone())?;
        let adjusted_index = normalize_array_index_operand(index_expr, COBOL_ARRAY_INDEXING);
        expr = Expression::new(ExprKind::Index {
            object: Box::new(expr),
            index: Box::new(adjusted_index),
            null_safe: false });

        for extra in expr_children.iter().skip(1) {
            let extra_idx = walk_expression((*extra).clone())?;
            let adjusted = normalize_array_index_operand(extra_idx, COBOL_ARRAY_INDEXING);
            expr = Expression::new(ExprKind::Index {
                object: Box::new(expr),
                index: Box::new(adjusted),
                null_safe: false });
        }
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
        _ => Err(format!(
            "COBOL walker: unhandled literal rule {:?}",
            inner.as_rule()
        )) }
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
        _ => Ok(Expression::bool(false)) }
}

fn parse_number_literal(s: &str) -> Expression {
    let trimmed = s.trim();
    let normalized = if trimmed.contains(',') && !trimmed.contains('.') {
        trimmed.replace(',', ".")
    } else {
        trimmed.to_string()
    };

    if normalized.contains('.') {
        match normalized.parse::<f64>() {
            Ok(f) => Expression::float(f),
            Err(_) => Expression::int(0) }
    } else {
        match normalized.parse::<i64>() {
            Ok(n) => Expression::int(n),
            Err(_) => Expression::int(0) }
    }
}

fn walk_string_literal(pair: &Pair<Rule>) -> Result<Expression, String> {
    let raw = pair.as_str();
    // Strip surrounding quotes (either " or ')
    if raw.len() >= 2 {
        let inner = &raw[1..raw.len() - 1];
        Ok(Expression::string(inner))
    } else {
        Ok(Expression::string(""))
    }
}

fn string_literal_value(pair: &Pair<Rule>) -> String {
    let raw = pair.as_str();
    if raw.len() >= 2 {
        raw[1..raw.len() - 1].to_string()
    } else {
        String::new()
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Helper functions
// ════════════════════════════════════════════════════════════════════════════

/// Create a binary expression.
fn cobol_name_key(name: &str) -> String {
    name.trim().to_ascii_uppercase()
}

fn cobol_data_item_level(pair: &Pair<Rule>) -> u32 {
    match pair.as_rule() {
        Rule::data_item => pair
            .clone()
            .into_inner()
            .find_map(|child| match child.as_rule() {
                Rule::regular_data_item => Some(cobol_data_item_level(&child)),
                _ => None })
            .unwrap_or(0),
        Rule::regular_data_item => pair
            .clone()
            .into_inner()
            .find(|child| child.as_rule() == Rule::level_number)
            .and_then(|child| child.as_str().trim().parse::<u32>().ok())
            .unwrap_or(0),
        _ => 0 }
}

/// Create a binary expression.
fn binary(op: BinOp, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right) })
}

/// Negate an expression (wrap in NOT).
fn negate_expr(expr: Expression) -> Expression {
    Expression::new(ExprKind::Unary {
        op: UnaryOp::Not,
        expr: Box::new(expr) })
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
