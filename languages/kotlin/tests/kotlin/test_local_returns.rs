use crate::helpers::run_prints;

#[test]
fn test_local_return_from_for_each() {
    let out = run_prints(r#"
        fun main() {
            val out = mutableListOf<Int>()
            listOf(1, 2, 3, 4).forEach {
                if (it == 3) return@forEach
                out.add(it)
            }
            println(out.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,2,4"]);
}

#[test]
fn test_local_return_from_while_like_loop() {
    let out = run_prints(r#"
        fun main() {
            var i = 0
            val out = StringBuilder()
            while (i < 5) {
                if (i == 3) { i += 1; continue }
                out.append(i)
                i += 1
            }
            println(out.toString())
        }
    "#);
    assert_eq!(out, &["0124"]);
}

#[test]
fn test_local_return_lambda_named() {
    let out = run_prints(r#"
        fun main() {
            fun run(items: List<Int>): String {
                val r = StringBuilder()
                items.forEach loop@{
                    if (it == 2) return@loop
                    r.append(it)
                }
                return r.toString()
            }
            println(run(listOf(1, 2, 3)))
        }
    "#);
    assert_eq!(out, &["13"]);
}

#[test]
fn test_local_return_non_local_from_inline() {
    let out = run_prints(r#"
        fun sum(values: List<Int>): Int {
            var s = 0
            values.forEach {
                if (it < 0) return 0
                s += it
            }
            return s
        }
        fun main() {
            println(sum(listOf(1, 2, 3)))
            println(sum(listOf(-1, 2)))
        }
    "#);
    assert_eq!(out, &["6", "0"]);
}

#[test]
fn test_local_return_named_label_in_run() {
    let out = run_prints(r#"
        fun main() {
            val result = run label@{
                listOf(1, 2).forEach {
                    if (it == 1) return@label "first"
                }
                "none"
            }
            println(result)
        }
    "#);
    assert_eq!(out, &["first"]);
}

#[test]
fn test_local_return_in_apply() {
    let out = run_prints(r#"
        fun main() {
            val result = apply("") {
                this += "x"
                return@apply this
            }
            println(result)
        }
    "#);
    assert_eq!(out, &["x"]);
}

#[test]
fn test_local_return_in_with() {
    let out = run_prints(r#"
        fun main() {
            val value = with(0) {
                if (this < 0) return@with 0
                this + 1
            }
            println(value)
        }
    "#);
    assert_eq!(out, &["1"]);
}

#[test]
fn test_local_return_in_let_binding() {
    let out = run_prints(r#"
        fun main() {
            val v = "x".let {
                if (it.isEmpty()) return@let "empty"
                "val=" + it
            }
            println(v)
        }
    "#);
    assert_eq!(out, &["val=x"]);
}

#[test]
fn test_local_return_in_also() {
    let out = run_prints(r#"
        fun main() {
            val n = 4
            val out = n.also {
                if (it < 0) return
            }
            println(out)
        }
    "#);
    assert_eq!(out, &["4"]);
}

#[test]
fn test_local_return_in_take_if() {
    let out = run_prints(r#"
        fun main() {
            val v = 3
            val out = v.takeIf { it > 1 } ?: 0
            println(out)
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_local_return_in_take_unless() {
    let out = run_prints(r#"
        fun main() {
            val v = 0
            val out = v.takeUnless { it > 1 } ?: 9
            println(out)
        }
    "#);
    assert_eq!(out, &["9"]);
}

#[test]
fn test_local_return_in_anonymous_function() {
    let out = run_prints(r#"
        fun main() {
            val f: (Int) -> Int = fun(x: Int): Int {
                if (x == 0) return 5
                return x
            }
            println(f(0))
            println(f(2))
        }
    "#);
    assert_eq!(out, &["5", "2"]);
}

#[test]
fn test_local_return_from_try() {
    let out = run_prints(r#"
        fun main() {
            val out = run {
                try {
                    throw Exception("x")
                } catch (e: Exception) {
                    return@run "err"
                }
                "ok"
            }
            println(out)
        }
    "#);
    assert_eq!(out, &["err"]);
}

#[test]
fn test_local_return_from_when() {
    let out = run_prints(r#"
        fun status(v: Int): String {
            return when (v) {
                in 1..3 -> "small"
                else -> return "other"
            }
        }
        fun main() {
            println(status(2))
            println(status(10))
        }
    "#);
    assert_eq!(out, &["small", "other"]);
}

#[test]
fn test_local_return_in_builder() {
    let out = run_prints(r#"
        fun main() {
            val result = StringBuilder().apply {
                append("a")
                if (length == 0) return
                append("b")
            }
            println(result.toString())
        }
    "#);
    assert_eq!(out, &["ab"]);
}

#[test]
fn test_local_return_breakout_loop() {
    let out = run_prints(r#"
        fun main() {
            val out = StringBuilder()
            for (i in 1..5) {
                if (i == 4) break
                out.append(i)
            }
            println(out.toString())
        }
    "#);
    assert_eq!(out, &["123"]);
}

#[test]
fn test_local_return_continue_loop() {
    let out = run_prints(r#"
        fun main() {
            val out = StringBuilder()
            for (i in 1..4) {
                if (i == 2) continue
                out.append(i)
            }
            println(out.toString())
        }
    "#);
    assert_eq!(out, &["134"]);
}

#[test]
fn test_local_return_nested_for_each() {
    let out = run_prints(r#"
        fun main() {
            val result = StringBuilder()
            outer@ for (r in 1..2) {
                inner@ for (c in 1..3) {
                    if (r == 1 && c == 2) continue@outer
                    result.append(r).append(c)
                }
            }
            println(result.toString())
        }
    "#);
    assert_eq!(out, &["11123131"]);
}

#[test]
fn test_local_return_in_repeat() {
    let out = run_prints(r#"
        fun main() {
            val out = StringBuilder()
            repeat(4) { index ->
                if (index == 2) return@repeat
                out.append(index)
            }
            println(out.toString())
        }
    "#);
    assert_eq!(out, &["013"]);
}

#[test]
fn test_local_return_in_each_indexed() {
    let out = run_prints(r#"
        fun main() {
            val out = StringBuilder()
            listOf(1, 2, 3).forEachIndexed { index, value ->
                if (index == 1) return@forEachIndexed
                out.append(value)
            }
            println(out.toString())
        }
    "#);
    assert_eq!(out, &["13"]);
}

#[test]
fn test_local_return_in_map_filter() {
    let out = run_prints(r#"
        fun main() {
            val filtered = listOf(1, 2, 3).filter {
                if (it % 2 == 0) return@filter false
                true
            }
            println(filtered.joinToString())
        }
    "#);
    assert_eq!(out, &["1, 3"]);
}

#[test]
fn test_local_return_in_any() {
    let out = run_prints(r#"
        fun main() {
            val anyEven = listOf(1, 3, 4).any {
                if (it == 2) return@any false
                it % 2 == 0
            }
            println(anyEven)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_local_return_in_all() {
    let out = run_prints(r#"
        fun main() {
            val allPositive = listOf(1, 2, -1).all {
                if (it < 0) return@all false
                it > 0
            }
            println(allPositive)
        }
    "#);
    assert_eq!(out, &["false"]);
}

#[test]
fn test_local_return_in_fold() {
    let out = run_prints(r#"
        fun main() {
            val total = listOf(1, 2, 3).fold(0) { acc, value ->
                if (value == 2) return@fold acc
                acc + value
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["4"]);
}

#[test]
fn test_local_return_nonlocal_for_each() {
    let out = run_prints(r#"
        fun search(values: List<Int>): String {
            for (v in values) {
                if (v == 3) return "found"
            }
            return "none"
        }
        fun main() {
            println(search(listOf(1, 2, 3)))
            println(search(listOf(1, 2)))
        }
    "#);
    assert_eq!(out, &["found", "none"]);
}

#[test]
fn test_local_return_with_map_update() {
    let out = run_prints(r#"
        fun main() {
            val out = mutableMapOf<Int, Int>()
            for (i in 1..4) {
                if (i == 3) continue
                out[i] = i
            }
            println(out.size)
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_local_return_empty_block() {
    let out = run_prints(r#"
        fun main() {
            val result = run {
                val v = 0
                if (v == 0) return@run 0
                1
            }
            println(result)
        }
    "#);
    assert_eq!(out, &["0"]);
}

#[test]
fn test_local_return_in_sequence_map() {
    let out = run_prints(r#"
        fun main() {
            val out = sequenceOf(1, 2, 3).map {
                if (it == 1) return@map 10
                it
            }.toList().joinToString(",")
            println(out)
        }
    "#);
    assert_eq!(out, &["10,2,3"]);
}

#[test]
fn test_local_return_in_run_block_string() {
    let out = run_prints(r#"
        fun main() {
            val text = "abc"
            val out = run {
                if (text.isEmpty()) return@run "empty"
                text
            }
            println(out)
        }
    "#);
    assert_eq!(out, &["abc"]);
}

#[test]
fn test_local_return_in_when_guard() {
    let out = run_prints(r#"
        fun main() {
            val value = 5
            val label = when {
                value > 10 -> "high"
                value < 0 -> "low"
                else -> run { value.toString() }
            }
            println(label)
        }
    "#);
    assert_eq!(out, &["5"]);
}

#[test]
fn test_local_return_in_list_any_match() {
    let out = run_prints(r#"
        fun main() {
            val hasLarge = listOf(1, 2, 8, 9).any {
                if (it >= 8) return@any true
                false
            }
            println(hasLarge)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_local_return_in_flat_map() {
    let out = run_prints(r#"
        fun main() {
            val data = listOf(1, 2, 3).flatMap {
                if (it == 2) return@flatMap listOf<Int>()
                listOf(it)
            }
            println(data.joinToString("/"))
        }
    "#);
    assert_eq!(out, &["1/3"]);
}

#[test]
fn test_local_return_in_map_get() {
    let out = run_prints(r#"
        fun main() {
            val map = mapOf("a" to 1, "b" to 2)
            val value = map["c"] ?: run {
                if (map.isEmpty()) 0 else 99
            }
            println(value)
        }
    "#);
    assert_eq!(out, &["99"]);
}

#[test]
fn test_local_return_lambda_with_label() {
    let out = run_prints(r#"
        fun main() {
            val out = listOf(1, 2, 3).filterIndexed { index, value ->
                if (index == 1) return@filterIndexed false
                value % 2 == 1
            }
            println(out.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,3"]);
}

#[test]
fn test_local_return_in_try_finally() {
    let out = run_prints(r#"
        fun main() {
            val out = run {
                try {
                    return@run "ok"
                } finally {
                    println("fin")
                }
            }
            println(out)
        }
    "#);
    assert_eq!(out, &["fin", "ok"]);
}

#[test]
fn test_local_return_in_repeat_with_indices() {
    let out = run_prints(r#"
        fun main() {
            val out = StringBuilder()
            repeat(3) {
                if (it == 1) return@repeat
                out.append(it)
            }
            println(out.toString())
        }
    "#);
    assert_eq!(out, &["02"]);
}

#[test]
fn test_local_return_in_while_conditional() {
    let out = run_prints(r#"
        fun main() {
            var i = 0
            while (i < 4) {
                i += 1
                if (i == 2) continue
                if (i == 4) break
                println(i)
            }
        }
    "#);
    assert_eq!(out, &["1", "3"]);
}
