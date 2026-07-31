use crate::helpers::run_prints;

#[test]
fn test_logical_and_skips_right_when_false() {
    let out = run_prints(r#"
        fun main() {
            var log = ""
            fun right(): Boolean { log += "right"; return true }
            println(false && right())
            println(log)
        }
    "#);
    assert_eq!(out, &["false", ""]);
}

#[test]
fn test_logical_or_skips_right_when_left_true() {
    let out = run_prints(r#"
        fun main() {
            var log = ""
            fun right(): Boolean { log += "right"; return false }
            println(true || right())
            println(log)
        }
    "#);
    assert_eq!(out, &["true", ""]);
}

#[test]
fn test_logical_or_evaluates_right_when_left_false() {
    let out = run_prints(r#"
        fun main() {
            var log = ""
            fun right(): Boolean { log += "right"; return false }
            println(false || right())
            println(log)
        }
    "#);
    assert_eq!(out, &["false", "right"]);
}

#[test]
fn test_logical_and_evaluates_right_when_left_true() {
    let out = run_prints(r#"
        fun main() {
            var log = ""
            fun right(): Boolean { log += "right"; return false }
            println(true && right())
            println(log)
        }
    "#);
    assert_eq!(out, &["false", "right"]);
}

#[test]
fn test_chained_and_short_circuit_order() {
    let out = run_prints(r#"
        fun main() {
            var log = ""
            fun a(): Boolean { log += "a"; return false }
            fun b(): Boolean { log += "b"; return false }
            fun c(): Boolean { log += "c"; return true }
            println(a() && b() && c())
            println(log)
        }
    "#);
    assert_eq!(out, &["false", "a"]);
}

#[test]
fn test_chained_or_short_circuit_order() {
    let out = run_prints(r#"
        fun main() {
            var log = ""
            fun a(): Boolean { log += "a"; return true }
            fun b(): Boolean { log += "b"; return false }
            fun c(): Boolean { log += "c"; return true }
            println(a() || b() || c())
            println(log)
        }
    "#);
    assert_eq!(out, &["true", "a"]);
}

#[test]
fn test_mix_or_before_and_respects_grouping() {
    let out = run_prints(r#"
        fun main() {
            var log = ""
            fun l(v: String): Boolean { log += v; return v == "go" }
            println((l("go") || l("skip")) && l("and"))
            println(log)
        }
    "#);
    assert_eq!(out, &["true", "go"]);
}

#[test]
fn test_mix_and_before_or_is_right_associative_like_call_graph() {
    let out = run_prints(r#"
        fun main() {
            var log = ""
            fun l(v: String): Boolean {
                log += v
                return v == "go"
            }
            println(l("go") && l("and") || l("tail"))
            println(log)
        }
    "#);
    assert_eq!(out, &["true", "goand"]);
}

#[test]
fn test_bitwise_and_is_not_boolean_short_circuit() {
    let out = run_prints(r#"
        fun main() {
            println(1 and 0)
            println(1 or 2)
        }
    "#);
    assert_eq!(out, &["0", "3"]);
}

#[test]
fn test_safe_call_does_not_follow_false_branch() {
    let out = run_prints(r#"
        fun main() {
            var log = ""
            fun sideEffect(v: String): String { log += v; return v }
            val x: String? = null
            println((x == null) && (sideEffect("bad") == "bad") )
            println(log)
            val y: String? = "ok"
            println((y != null) || (sideEffect("bad") == "bad") )
            println(log)
        }
    "#);
    assert_eq!(out, &["true", "", "true", ""]);
}

#[test]
fn test_when_with_guard_uses_short_circuit_or() {
    let out = run_prints(r#"
        fun main() {
            val value = 5
            var log = ""
            val out = when (value) {
                1, 2, 3 -> { log += "a"; 1 }
                in 4..10 -> { log += "b"; 2 }
                else -> { log += "c"; 3 }
            }
            println(out)
            println(log)
        }
    "#);
    assert_eq!(out, &["2", "b"]);
}

#[test]
fn test_boolean_chain_with_comparison_short_circuit() {
    let out = run_prints(r#"
        fun main() {
            var log = 0
            fun left(): Int { log += 1; return 1 }
            fun right(): Int { log += 10; return 2 }
            val v = (left() == 1) && (right() == 3)
            println(v)
            println(log)
        }
    "#);
    assert_eq!(out, &["false", "1"]);
}

#[test]
fn test_or_chain_with_all_rhs_calls() {
    let out = run_prints(r#"
        fun main() {
            var log = ""
            fun a(): Boolean { log += "a"; return false }
            fun b(): Boolean { log += "b"; return false }
            fun c(): Boolean { log += "c"; return true }
            println(a() || b() || c())
            println(log)
        }
    "#);
    assert_eq!(out, &["true", "abc"]);
}

#[test]
fn test_and_chain_with_final_false() {
    let out = run_prints(r#"
        fun main() {
            var log = ""
            fun a(): Boolean { log += "a"; return true }
            fun b(): Boolean { log += "b"; return true }
            fun c(): Boolean { log += "c"; return false }
            println(a() && b() && c())
            println(log)
        }
    "#);
    assert_eq!(out, &["false", "abc"]);
}

#[test]
fn test_guarded_if_skips_false_branch_side_effect() {
    let out = run_prints(r#"
        fun main() {
            var log = 0
            val a = true
            if (a && false) {
                log += 1
            } else {
                log += 2
            }
            println(log)
        }
    "#);
    assert_eq!(out, &["2"]);
}

#[test]
fn test_short_circuit_with_nullable_lhs() {
    let out = run_prints(r#"
        fun main() {
            val text: String? = null
            println(text != null && text.isNotEmpty())
        }
    "#);
    assert_eq!(out, &["false"]);
}

#[test]
fn test_short_circuit_with_nullable_rhs_not_needed() {
    let out = run_prints(r#"
        fun main() {
            var log = ""
            val text: String? = ""
            fun check(): Boolean {
                log += "check"
                return true
            }
            println(text != null && check())
            println(log)
        }
    "#);
    assert_eq!(out, &["true", "check"]);
}

#[test]
fn test_or_prefers_true_without_evaluating_rhs() {
    let out = run_prints(r#"
        fun main() {
            var log = ""
            fun rhs(): Boolean {
                log += "rhs"
                return false
            }
            println((5 > 2) || rhs())
            println(log)
        }
    "#);
    assert_eq!(out, &["true", ""]);
}

#[test]
fn test_iffalse_and_rhs_not_called_even_with_other_predicates() {
    let out = run_prints(r#"
        fun main() {
            var log = ""
            fun rhs(): Boolean {
                log += "rhs"
                return true
            }
            println((0 > 1) && rhs())
            println(log)
        }
    "#);
    assert_eq!(out, &["false", ""]);
}

#[test]
fn test_short_circuit_preserves_side_effect_order_for_or() {
    let out = run_prints(r#"
        fun main() {
            var log = ""
            fun a(): Boolean { log += "1"; return false }
            fun b(): Boolean { log += "2"; return true }
            println(a() || b())
            println(log)
        }
    "#);
    assert_eq!(out, &["true", "12"]);
}

#[test]
fn test_short_circuit_preserves_side_effect_order_for_and() {
    let out = run_prints(r#"
        fun main() {
            var log = ""
            fun a(): Boolean { log += "1"; return true }
            fun b(): Boolean { log += "2"; return false }
            println(a() && b())
            println(log)
        }
    "#);
    assert_eq!(out, &["false", "12"]);
}

#[test]
fn test_truth_table_basics_and_or() {
    let out = run_prints(r#"
        fun main() {
            println(true || false)
            println(false || false)
            println(true && true)
            println(false && true)
        }
    "#);
    assert_eq!(out, &["true", "false", "true", "false"]);
}
