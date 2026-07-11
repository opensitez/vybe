#[macro_use]
mod helpers;

// ── Foundations: types, conditionals, loops, strings, tables ───────────────
mod test_assignment;
mod test_basics;
mod test_coercion;
mod test_control_flow;
mod test_core_syntax;
mod test_iteration;
mod test_literals;
mod test_loops_for_generic;
mod test_loops_for_numeric;
mod test_loops_repeat_until;
mod test_loops_while;
mod test_operators;
mod test_strings;
mod test_table_constructors;
mod test_tables;
mod test_tables_arrays;
mod test_tables_custom_iterators;
mod test_tables_hash;
mod test_tables_mixed;
mod test_tables_weak_keys;
mod test_tables_weak_kv;
mod test_tables_weak_values;
mod test_truthiness;

// ── Core architecture: functions, scope, varargs, patterns, errors ─────────
mod test_closures;
mod test_closures_complex;
mod test_closures_ext;
mod test_closures_nested;
mod test_closures_upvalues;
mod test_error_handling_complex;
mod test_error_handling_ext;
mod test_error_handling_pcall;
mod test_error_handling_xpcall;
mod test_error_levels;
mod test_errors;
mod test_functions;
mod test_functions_dump;
mod test_functions_load;
mod test_functions_multiple_returns;
mod test_functions_pcall_xpcall;
mod test_functions_tailcalls;
mod test_functions_vararg;
mod test_functions_vararg_ext;
mod test_globals;
mod test_metatables_proxy_tables;
mod test_oop;
mod test_scoping;
mod test_scoping_blocks;
mod test_string_gsub_ext;
mod test_string_matching;
mod test_string_patterns;
mod test_string_patterns_complex;
mod test_tables_metatables_oop;
mod test_vararg;

// ── Metaprogramming: metatables, modules, coroutines, math, bitwise ───────
mod test_bitwise;
mod test_bitwise_logic;
mod test_bitwise_shifts;
mod test_coroutine_ext;
mod test_coroutine_wrappers;
mod test_coroutines;
mod test_coroutines_advanced;
mod test_coroutines_basics;
mod test_coroutines_errors;
mod test_coroutines_wrap;
mod test_coroutines_yield_resume;
mod test_environment;
mod test_environment_lua52;
mod test_environments_ext;
mod test_lexical_environments_advanced;
mod test_literals_hex;
mod test_literals_long_strings;
mod test_load;
mod test_math_integers;
mod test_math_library;
mod test_math_log_exp;
mod test_math_misc;
mod test_math_random;
mod test_math_random_ext;
mod test_math_trig;
mod test_metamethods;
mod test_metatables;
mod test_metatables_call;
mod test_metatables_concat;
mod test_metatables_index;
mod test_metatables_len;
mod test_metatables_math;
mod test_metatables_newindex;
mod test_modules_complex;
mod test_modules_package_path;
mod test_modules_require;
mod test_operators_arithmetic;
mod test_operators_concat;
mod test_operators_length;
mod test_operators_logical;
mod test_operators_relational;
mod test_package;
mod test_package_ext;
mod test_tables_metatables_ext;
mod test_type_checks;
mod test_type_coercion;

// ── Standard libraries & integration ───────────────────────────────────────
mod test_collectgarbage;
mod test_debug_hooks;
mod test_debug_introspection;
mod test_debug_library;
mod test_debug_locals;
mod test_debug_upvalues;
mod test_garbage_collection_advanced;
mod test_garbage_collection_api;
mod test_garbage_collection_finalizers;
mod test_goto;
mod test_goto_labels;
mod test_io_file_handles;
mod test_io_implicit;
mod test_io_library;
mod test_io_lines_ext;
mod test_io_popen;
mod test_io_read_ext;
mod test_nan_and_inf;
mod test_os_library;
mod test_os_misc;
mod test_os_time_date;
mod test_programs;
mod test_string_char_byte;
mod test_string_find;
mod test_string_format;
mod test_string_format_ext;
mod test_string_pack;
mod test_string_pack_unpack_ext;
mod test_string_rep_rev;
mod test_string_sub_len;
mod test_table_concat;
mod test_table_insert_remove;
mod test_tables_mixed_ext;
mod test_tables_move;
mod test_tables_pack_unpack;
mod test_tables_sort_custom;
mod test_tail_calls;
mod test_utf8;
mod test_utf8_ext;
mod test_utf8_iteration;
mod test_utf8_validation;
mod test_weak_tables;

mod test_select_builtin;
mod test_tostring_tonumber;
mod test_string_gmatch;
mod test_raw_access;
mod test_table_unpack_range;
mod test_multiple_return_adjustment;
mod test_next_traversal;
mod test_pcall_error_objects;
mod test_ipairs_pairs_edge;
mod test_metamethods_bitwise;
mod test_string_format_specifiers;
mod test_metamethods_comparison;
mod test_generic_for_protocol;
mod test_string_patterns_frontier_balanced;
mod test_xpcall_handler;
mod test_loops_numeric_edge;
mod test_load_custom_env;
mod test_math_constants_rounding;
mod test_do_blocks;
mod test_string_gsub_replacements;
mod test_coroutine_patterns;
mod test_module_tables;
mod test_metatables_index_chains;
mod test_os_date_time;
mod test_vararg_advanced;
mod test_functional_patterns;
mod test_metatable_protection;
mod test_oop_metatable_patterns;

mod test_operators_logical_advanced;
mod test_scope_variable_lifetime;
mod test_math_trig_advanced;
mod test_string_gsub_count;
mod test_table_constructors_advanced;
mod test_goto_advanced;
mod test_pcall_patterns;
mod test_coroutines_state_machine;
mod test_string_matching_captures;
mod test_metatables_arithmetic;
mod test_debug_library_basics;
mod test_package_system;

mod test_coroutines_nested_yield;
mod test_string_patterns_captures_advanced;
mod test_utf8_char_advanced;
mod test_math_random_seeding;
mod test_error_handling_xpcall_nested;
mod test_lexical_scoping_advanced;
mod test_vararg_functions_advanced;
mod test_table_iteration_order;
mod test_metamethods_relational;
mod test_env_lexical_binding;

mod test_math_library_extended;
mod test_string_patterns_extended;
mod test_table_library_extended;
mod test_coroutines_extended;
mod test_metatables_extended;
mod test_language_semantics_extended;

mod test_metatables_fallback_inheritance;
mod test_pcall_nested_xpcall_scenarios;
mod test_goto_nested_scoping;
mod test_math_advanced_log_exp;
mod test_package_path_searching;
mod test_weak_tables_advanced;
mod test_operators_bitwise_advanced;
mod test_do_blocks_scoping_advanced;
mod test_string_format_advanced;

mod test_math_exhaustive;
mod test_string_patterns_exhaustive;

mod test_operators_exhaustive;
mod test_control_flow_exhaustive;
mod test_types_exhaustive;

mod test_metatables_exhaustive;
mod test_string_exhaustive;
mod test_table_exhaustive;
mod test_coroutines_exhaustive;
mod test_scoping_exhaustive;
mod test_math_advanced_exhaustive;

mod test_operators_heavy;
mod test_control_flow_heavy;

mod test_operators_super;
mod test_control_flow_super;

mod test_types_super;
mod test_string_super;
mod test_table_super;











