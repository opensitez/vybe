use crate::helpers::run_prints;

#[test]
fn test_template_basic_identifier() {
    let out = run_prints(
        r#"
        fun main() {
            val name = "k"
            println("hello $name")
        }
    "#,
    );
    assert_eq!(out, &["hello k"]);
}

#[test]
fn test_template_expression_sum() {
    let out = run_prints(
        r#"
        fun main() {
            val a = 2
            val b = 3
            println("sum=${a + b}")
        }
    "#,
    );
    assert_eq!(out, &["sum=5"]);
}

#[test]
fn test_template_function_call() {
    let out = run_prints(
        r#"
        fun fmt(v: Int): String = "hash" + v.toString()
        fun main() {
            println("v=${fmt(4)}")
        }
    "#,
    );
    assert_eq!(out, &["v=hash4"]);
}

#[test]
fn test_template_method_call() {
    let out = run_prints(
        r#"
        fun main() {
            println("upper=${"ab".uppercase()}")
        }
    "#,
    );
    assert_eq!(out, &["upper=AB"]);
}

#[test]
fn test_template_if_expression() {
    let out = run_prints(
        r#"
        fun main() {
            val ok = true
            println("status=${if (ok) "ok" else "bad"}")
        }
    "#,
    );
    assert_eq!(out, &["status=ok"]);
}

#[test]
fn test_template_when_expression() {
    let out = run_prints(
        r#"
        fun main() {
            val x = 2
            println("v=${when(x) { 1 -> "a" 2 -> "b" else -> "c" }}")
        }
    "#,
    );
    assert_eq!(out, &["v=b"]);
}

#[test]
fn test_template_string_concat_mix() {
    let out = run_prints(
        r#"
        fun main() {
            val p = "a"
            val q = "b"
            println("$p+$q=${p + q}")
        }
    "#,
    );
    assert_eq!(out, &["a+b=ab"]);
}

#[test]
fn test_template_with_array_size() {
    let out = run_prints(
        r#"
        fun main() {
            val items = listOf(1, 2, 3)
            println("size=${items.size}")
        }
    "#,
    );
    assert_eq!(out, &["size=3"]);
}

#[test]
fn test_template_nullable_value_present() {
    let out = run_prints(
        r#"
        fun main() {
            val v: String? = "x"
            println("value=${v}")
        }
    "#,
    );
    assert_eq!(out, &["value=x"]);
}

#[test]
fn test_template_nullable_value_null() {
    let out = run_prints(
        r#"
        fun main() {
            val v: String? = null
            println("value=${v}")
        }
    "#,
    );
    assert_eq!(out, &["value=null"]);
}

#[test]
fn test_template_property_accessor() {
    let out = run_prints(
        r#"
        class User(val name: String)
        fun main() {
            val u = User("a")
            println("user=${u.name}")
        }
    "#,
    );
    assert_eq!(out, &["user=a"]);
}

#[test]
fn test_template_nested_calls() {
    let out = run_prints(
        r#"
        fun main() {
            println("len=${"abc".length + "de".length}")
        }
    "#,
    );
    assert_eq!(out, &["len=5"]);
}

#[test]
fn test_template_raw_like_quoted() {
    let out = run_prints(
        r#"
        fun main() {
            println("path=/tmp/${"a"}")
        }
    "#,
    );
    assert_eq!(out, &["path=/tmp/a"]);
}

#[test]
fn test_template_escape_dollar_sign() {
    let out = run_prints(
        r#"
        fun main() {
            println("price=\$10")
        }
    "#,
    );
    assert_eq!(out, &["price=$10"]);
}

#[test]
fn test_template_in_array_join() {
    let out = run_prints(
        r#"
        fun main() {
            val v = listOf("x", "y", "z")
            println("joined=${v.joinToString(",")}")
        }
    "#,
    );
    assert_eq!(out, &["joined=x,y,z"]);
}

#[test]
fn test_template_with_boolean() {
    let out = run_prints(
        r#"
        fun main() {
            val ok = false
            println("flag=${ok}")
        }
    "#,
    );
    assert_eq!(out, &["flag=false"]);
}

#[test]
fn test_template_with_char() {
    let out = run_prints(
        r#"
        fun main() {
            val c: Char = 'q'
            println("char=$c")
        }
    "#,
    );
    assert_eq!(out, &["char=q"]);
}

#[test]
fn test_template_with_double() {
    let out = run_prints(
        r#"
        fun main() {
            val d = 1.5
            println("double=$d")
        }
    "#,
    );
    assert_eq!(out, &["double=1.5"]);
}

#[test]
fn test_template_with_math() {
    let out = run_prints(
        r#"
        fun main() {
            println("pi=${kotlin.math.PI}")
        }
    "#,
    );
    assert_eq!(out, &["pi=3.141592653589793"]);
}

#[test]
fn test_template_with_indexed_access() {
    let out = run_prints(
        r#"
        fun main() {
            val s = "abc"
            println("third=${s[2]}")
        }
    "#,
    );
    assert_eq!(out, &["third=c"]);
}

#[test]
fn test_template_with_let_chain() {
    let out = run_prints(
        r#"
        fun main() {
            val out = listOf(1, 2, 3).let { it.size }
            println("let=$out")
        }
    "#,
    );
    assert_eq!(out, &["let=3"]);
}

#[test]
fn test_template_with_map_lookup() {
    let out = run_prints(
        r#"
        fun main() {
            val m = mapOf("a" to 7)
            println("lookup=${m["a"]}")
        }
    "#,
    );
    assert_eq!(out, &["lookup=7"]);
}

#[test]
fn test_template_multiline_concat() {
    let out = run_prints(
        r#"
        fun main() {
            val a = 1
            val b = 2
            println("a=${a}")
            println("sum=${a + b}")
        }
    "#,
    );
    assert_eq!(out, &["a=1", "sum=3"]);
}

#[test]
fn test_template_block_of_string() {
    let out = run_prints(
        r#"
        fun main() {
            val x = 3
            println("value=${if (x % 2 == 0) "even" else "odd"}")
        }
    "#,
    );
    assert_eq!(out, &["value=odd"]);
}

#[test]
fn test_template_with_nested_list() {
    let out = run_prints(
        r#"
        fun main() {
            val grid = listOf(listOf(1, 2), listOf(3, 4))
            println("rows=${grid.size} first=${grid[0][0]}")
        }
    "#,
    );
    assert_eq!(out, &["rows=2 first=1"]);
}

#[test]
fn test_template_empty_string() {
    let out = run_prints(
        r#"
        fun main() {
            val s = ""
            println("empty=${s.isEmpty()}")
        }
    "#,
    );
    assert_eq!(out, &["empty=true"]);
}

#[test]
fn test_template_range_join() {
    let out = run_prints(
        r#"
        fun main() {
            println("range=${1..3}")
        }
    "#,
    );
    assert_eq!(out, &["range=1..3"]);
}

#[test]
fn test_template_with_conditional_operator() {
    let out = run_prints(
        r#"
        fun main() {
            val n = -1
            println("isPositive=${if (n > 0) "yes" else "no"}")
        }
    "#,
    );
    assert_eq!(out, &["isPositive=no"]);
}

#[test]
fn test_template_custom_to_string() {
    let out = run_prints(
        r#"
        data class Point(val x: Int, val y: Int)
        fun main() {
            val p = Point(1, 2)
            println("point=$p")
        }
    "#,
    );
    assert_eq!(out, &["point=Point(x=1, y=2)"]);
}

#[test]
fn test_template_unicode_literal() {
    let out = run_prints(
        r#"
        fun main() {
            println("heart=\u2665")
        }
    "#,
    );
    assert_eq!(out, &["heart=♥"]);
}

#[test]
fn test_template_boolean_and_math_combo() {
    let out = run_prints(
        r#"
        fun main() {
            val ok = true
            val value = 4
            println("ok=${ok} total=${value * 2}")
        }
    "#,
    );
    assert_eq!(out, &["ok=true total=8"]);
}

#[test]
fn test_template_joiner_with_nulls() {
    let out = run_prints(
        r#"
        fun main() {
            val values: List<String?> = listOf("a", null, "b")
            println("join=${values.joinToString(",")}")
        }
    "#,
    );
    assert_eq!(out, &["join=a,null,b"]);
}

#[test]
fn test_template_no_placeholder() {
    let out = run_prints(
        r#"
        fun main() {
            println("no-template")
        }
    "#,
    );
    assert_eq!(out, &["no-template"]);
}

#[test]
fn test_template_double_quoted_brace() {
    let out = run_prints(
        r#"
        fun main() {
            val left = 1
            val right = 2
            println("${left}+${right}=${left + right}")
        }
    "#,
    );
    assert_eq!(out, &["1+2=3"]);
}
