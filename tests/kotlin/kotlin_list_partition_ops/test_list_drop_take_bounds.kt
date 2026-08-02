// vybe-test: kotlin/kotlin_list_partition_ops/test_list_drop_take_bounds
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_partition_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val src = listOf("a", "b", "c")
            __check((src.drop(1).toString()).toString(), "[b, c]")
            __check((src.take(2).toString()).toString(), "[a, b]")
        }
