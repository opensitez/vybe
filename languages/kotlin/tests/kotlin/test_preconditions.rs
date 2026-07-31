use crate::helpers::run_prints;

#[test]
fn test_require_true_does_not_throw() {
    let out = run_prints(r#"
        fun main() {
            require(true)
            println("ok")
        }
    "#);
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_require_false_throws_illegal_argument() {
    let out = run_prints(r#"
        fun main() {
            try {
                require(false)
                println("no")
            } catch (e: IllegalArgumentException) {
                println(e::class.simpleName)
            }
        }
    "#);
    assert_eq!(out, &["IllegalArgumentException"]);
}

#[test]
fn test_require_with_message() {
    let out = run_prints(r#"
        fun main() {
            try {
                require(false, { "bad" })
                println("no")
            } catch (e: IllegalArgumentException) {
                println(e.message)
            }
        }
    "#);
    assert_eq!(out, &["bad"]);
}

#[test]
fn test_require_not_null_non_null() {
    let out = run_prints(r#"
        fun main() {
            val value: String? = "ok"
            println(requireNotNull(value))
        }
    "#);
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_require_not_null_null_is_error() {
    let out = run_prints(r#"
        fun main() {
            try {
                val value: String? = null
                requireNotNull(value)
                println("no")
            } catch (e: IllegalArgumentException) {
                println("missing")
            }
        }
    "#);
    assert_eq!(out, &["missing"]);
}

#[test]
fn test_check_true() {
    let out = run_prints(r#"
        fun main() {
            check(1 + 1 == 2)
            println("pass")
        }
    "#);
    assert_eq!(out, &["pass"]);
}

#[test]
fn test_check_false_throws_illegal_state() {
    let out = run_prints(r#"
        fun main() {
            try {
                check(false)
                println("no")
            } catch (e: IllegalStateException) {
                println(e::class.simpleName)
            }
        }
    "#);
    assert_eq!(out, &["IllegalStateException"]);
}

#[test]
fn test_check_with_message() {
    let out = run_prints(r#"
        fun main() {
            try {
                check(false, { "state bad" })
                println("no")
            } catch (e: IllegalStateException) {
                println(e.message)
            }
        }
    "#);
    assert_eq!(out, &["state bad"]);
}

#[test]
fn test_error_function_throws() {
    let out = run_prints(r#"
        fun main() {
            try {
                error("fatal")
                println("no")
            } catch (e: IllegalStateException) {
                println(e.message)
            }
        }
    "#);
    assert_eq!(out, &["fatal"]);
}

#[test]
fn test_check_not_null_with_custom_name() {
    let out = run_prints(r#"
        fun main() {
            fun read(v: String?): String {
                return checkNotNull(v)
            }
            println(read("x"))
        }
    "#);
    assert_eq!(out, &["x"]);
}

#[test]
fn test_check_not_null_throws_for_null() {
    let out = run_prints(r#"
        fun main() {
            try {
                checkNotNull<String>(null)
                println("no")
            } catch (e: IllegalStateException) {
                println("thrown")
            }
        }
    "#);
    assert_eq!(out, &["thrown"]);
}

#[test]
fn test_require_orchestrates_guard_clause() {
    let out = run_prints(r#"
        fun clamp(value: Int): Int {
            require(value in 1..10)
            return value
        }

        fun main() {
            println(clamp(7))
            try {
                println(clamp(42))
            } catch (e: IllegalArgumentException) {
                println("bad")
            }
        }
    "#);
    assert_eq!(out, &["7", "bad"]);
}

#[test]
fn test_check_in_variant_logic() {
    let out = run_prints(r#"
        fun parseAge(age: Int?): Int {
            checkNotNull(age)
            check(age >= 0)
            return age
        }

        fun main() {
            try {
                println(parseAge(-1))
            } catch (e: IllegalStateException) {
                println("state")
            }
        }
    "#);
    assert_eq!(out, &["state"]);
}

#[test]
fn test_require_uses_lazy_message() {
    let out = run_prints(r#"
        fun main() {
            val side = arrayOf(0)
            try {
                require(side[0] > 0, { side[0] = 1; "message" })
            } catch (e: IllegalArgumentException) {
                println(side[0])
                println(e.message)
            }
        }
    "#);
    assert_eq!(out, &["1", "message"]);
}

#[test]
fn test_precondition_chain_for_password() {
    let out = run_prints(r#"
        fun validate(password: String?) {
            requireNotNull(password)
            require(password.length >= 4, { "short" })
            require(password.any { it.isDigit() }, { "digit missing" })
        }

        fun main() {
            try {
                validate("a1")
                println("ok")
            } catch (e: IllegalArgumentException) {
                println(e.message)
            }
        }
    "#);
    assert_eq!(out, &["short"]);
}

#[test]
fn test_require_accepts_true_predicate() {
    let out = run_prints(r#"
        fun main() {
            require(true) { "nope" }
            println("ok")
        }
    "#);
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_error_before_return_is_caught() {
    let out = run_prints(r#"
        fun risky(v: Int): Int {
            require(v > 0)
            return v
        }

        fun main() {
            val value = try {
                risky(0)
            } catch (e: IllegalArgumentException) {
                -1
            }
            println(value)
        }
    "#);
    assert_eq!(out, &["-1"]);
}

#[test]
fn test_check_and_require_with_same_expression() {
    let out = run_prints(r#"
        fun main() {
            val v = 3
            require(v > 0)
            check(v == 3)
            println(v)
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_require_only_throws_not_for_nan() {
    let out = run_prints(r#"
        fun main() {
            try {
                require(Double.NaN.isFinite())
                println("finite")
            } catch (e: IllegalArgumentException) {
                println("bad")
            }
        }
    "#);
    assert_eq!(out, &["bad"]);
}

#[test]
fn test_check_not_null_message() {
    let out = run_prints(r#"
        fun main() {
            try {
                checkNotNull<Int>(null)
            } catch (e: IllegalStateException) {
                println("missing")
            }
        }
    "#);
    assert_eq!(out, &["missing"]);
}

#[test]
fn test_error_in_pipeline() {
    let out = run_prints(r#"
        fun parseInt(value: String): Int {
            return value.toIntOrNull() ?: error("invalid")
        }

        fun main() {
            val out = try {
                parseInt("x")
            } catch (e: IllegalStateException) {
                -1
            }
            println(out)
        }
    "#);
    assert_eq!(out, &["-1"]);
}

#[test]
fn test_guarded_list_restriction() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(1, 2, 3)
            require(values.isNotEmpty())
            check(values.size == 3)
            println(values.sum())
        }
    "#);
    assert_eq!(out, &["6"]);
}

#[test]
fn test_double_reject_and_accept() {
    let out = run_prints(r#"
        fun main() {
            for (value in listOf(-1, 1)) {
                val ok = try {
                    require(value > 0)
                    "yes"
                } catch (e: IllegalArgumentException) {
                    "no"
                }
                println(ok)
            }
        }
    "#);
    assert_eq!(out, &["no", "yes"]);
}

#[test]
fn test_require_with_nullable_subject() {
    let out = run_prints(r#"
        fun main() {
            val value: String? = null
            val out = try {
                requireNotNull(value)
                "ok"
            } catch (e: IllegalArgumentException) {
                "none"
            }
            println(out)
        }
    "#);
    assert_eq!(out, &["none"]);
}

#[test]
fn test_check_contract_with_boolean_expression() {
    let out = run_prints(r#"
        fun main() {
            val total = 5 + 6
            check(total == 11) { "sum mismatch" }
            println(total)
        }
    "#);
    assert_eq!(out, &["11"]);
}
