use crate::helpers::run_prints;

#[test]
fn test_fold_with_initial_value_and_indexing_behavior() {
    let out = run_prints(r#"
        fun main() {
            val nums = listOf(1, 2, 3)
            val total = nums.fold(10) { acc, n -> acc + n }
            val product = nums.fold(1) { acc, n -> acc * n }
            println(total)
            println(product)
        }
    "#);
    assert_eq!(out, &["16", "6"]);
}

#[test]
fn test_fold_right_from_end() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf("a", "b", "c")
            val out = values.foldRight("") { item, acc -> item + acc }
            println(out)
        }
    "#);
    assert_eq!(out, &["abc"]);
}

#[test]
fn test_reduce_vs_reduce_or_null() {
    let out = run_prints(r#"
        fun main() {
            val a = listOf(1, 2, 3, 4)
            val r = a.reduce { acc, value -> acc + value }
            println(r)
            val b = emptyList<Int>()
            println(b.reduceOrNull() ?: "empty")
        }
    "#);
    assert_eq!(out, &["10", "empty"]);
}

#[test]
fn test_reduce_right_on_strings() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf("x", "y", "z")
            val out = values.reduceRight { item, acc -> item + "," + acc }
            println(out)
        }
    "#);
    assert_eq!(out, &["x,y,z"]);
}

#[test]
fn test_fold_right_indexed() {
    let out = run_prints(r##"
        fun main() {
            val values = listOf(4, 5, 6)
            val out = values.foldRightIndexed("") { index, value, acc ->
                acc + value.toString() + "#" + index.toString() + ";"
            }
            println(out)
        }
    "##);
    assert_eq!(out, &["6#2;5#1;4#0;"]);
}

#[test]
fn test_reduce_or_null_singleton() {
    let out = run_prints(r#"
        fun main() {
            val one = listOf(42)
            println(one.reduceOrNull { a, b -> a + b })
        }
    "#);
    assert_eq!(out, &["42"]);
}

#[test]
fn test_sum_and_sum_by_key_like_projection() {
    let out = run_prints(r#"
        fun main() {
            val rows = listOf(
                Pair("a", 1),
                Pair("b", 2),
                Pair("a", 3)
            )
            println(rows.sumOf { it.second })
            println(rows.filter { it.first == "a" }.sumOf { it.second })
        }
    "#);
    assert_eq!(out, &["6", "4"]);
}

#[test]
fn test_running_fold_sequence() {
    let out = run_prints(r#"
        fun main() {
            val out = listOf(1, 2, 3).runningFold(0) { acc, n -> acc + n }.toList()
            println(out.joinToString(","))
        }
    "#);
    assert_eq!(out, &["0,1,3,6"]);
}

#[test]
fn test_running_reduce_sequence() {
    let out = run_prints(r#"
        fun main() {
            val out = listOf(2, 3, 4).runningReduce { acc, n -> acc * n }
            println(out.joinToString(","))
            val empty = emptyList<Int>()
            try {
                println(empty.runningReduce { a, b -> a + b }.joinToString(","))
            } catch (e: Exception) {
                println("err")
            }
        }
    "#);
    assert_eq!(out, &["2,6,24", "err"]);
}

#[test]
fn test_running_fold_with_indexed_progression() {
    let out = run_prints(r#"
        fun main() {
            val parts = listOf(1, 1, 2, 3).runningFoldIndexed(0) { index, acc, item ->
                acc + item + index
            }
            println(parts.joinToString(","))
        }
    "#);
    assert_eq!(out, &["0,1,3,6,10"]);
}

#[test]
fn test_fold_right_exception_when_empty() {
    let out = run_prints(r#"
        fun main() {
            val empty = emptyList<Int>()
            try {
                println(empty.reduceRight { a, b -> a + b })
            } catch (e: Exception) {
                println("no")
            }
        }
    "#);
    assert_eq!(out, &["no"]);
}

#[test]
fn test_fold_over_range_projection() {
    let out = run_prints(r#"
        fun main() {
            val projected = (1..4).fold(0) { acc, v -> acc * 10 + v }
            val right = (1..4).foldRight(100) { value, acc -> acc + value }
            println(projected)
            println(right)
        }
    "#);
    assert_eq!(out, &["1234", "110"]);
}

#[test]
fn test_fold_order_non_commutative() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf("a", "b", "c")
            val left = values.fold("") { acc, value -> "${'$'}acc${'$'}{value}" }
            val right = values.foldRight("") { value, acc -> "${'$'}value${'$'}{acc}" }
            println(left)
            println(right)
        }
    "#);
    assert_eq!(out, &["abc", "abc"]);
}

#[test]
fn test_counting_and_aggregation_shortcuts() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(2, 4, 6, 7, 9)
            println(values.count { it % 2 == 0 })
            println(values.sumOf { it })
            println(values.average())
        }
    "#);
    assert_eq!(out, &["2", "28", "5.6"]);
}
