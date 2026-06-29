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
mod test_operators;
mod test_strings;
mod test_table_constructors;
mod test_tables;
mod test_truthiness;

// ── Core architecture: functions, scope, varargs, patterns, errors ─────────
mod test_closures;
mod test_errors;
mod test_functions;
mod test_globals;
mod test_oop;
mod test_scoping;
mod test_string_patterns;
mod test_vararg;

// ── Metaprogramming: metatables, modules, coroutines, math, bitwise ───────
mod test_bitwise;
mod test_coroutines;
mod test_environment;
mod test_load;
mod test_math_library;
mod test_metamethods;
mod test_metatables;
mod test_package;
mod test_type_checks;

// ── Standard libraries & integration ───────────────────────────────────────
mod test_collectgarbage;
mod test_debug_library;
mod test_goto;
mod test_io_library;
mod test_nan_and_inf;
mod test_os_library;
mod test_programs;
mod test_string_format;
mod test_string_pack;
mod test_tail_calls;
mod test_utf8;
mod test_weak_tables;
