#[macro_use]
mod helpers;

// ── Foundations: types, conditionals, loops, strings, tables ───────────────
mod test_basics;
mod test_truthiness;
mod test_literals;
mod test_operators;
mod test_control_flow;
mod test_iteration;
mod test_strings;
mod test_tables;
mod test_table_constructors;
mod test_coercion;
mod test_assignment;
mod test_core_syntax;

// ── Core architecture: functions, scope, varargs, patterns, errors ─────────
mod test_functions;
mod test_closures;
mod test_scoping;
mod test_globals;
mod test_vararg;
mod test_string_patterns;
mod test_errors;
mod test_oop;

// ── Metaprogramming: metatables, modules, coroutines, math, bitwise ───────
mod test_metatables;
mod test_metamethods;
mod test_package;
mod test_environment;
mod test_coroutines;
mod test_math_library;
mod test_bitwise;
mod test_load;
mod test_type_checks;

// ── Standard libraries & integration ───────────────────────────────────────
mod test_string_format;
mod test_string_pack;
mod test_utf8;
mod test_goto;
mod test_programs;
mod test_nan_and_inf;
mod test_os_library;
mod test_collectgarbage;
mod test_debug_library;
mod test_io_library;
mod test_tail_calls;
mod test_weak_tables;
