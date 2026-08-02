// vybe-test: kotlin/kotlin_list_partition_ops/test_list_partition_by_predicate
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_partition_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val src = listOf(1, 2, 3, 4)
            val p = src.partition { it % 2 == 0 }
            __check((p.first.toString()).toString(), "[2, 4]")
            __check((p.second.toString()).toString(), "[1, 3]")
        }
