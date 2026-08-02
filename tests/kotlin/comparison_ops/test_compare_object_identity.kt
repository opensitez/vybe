// vybe-test: kotlin/comparison_ops/test_compare_object_identity
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

class Item
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Item()
            val b = Item()
            val c = a
            __check((a === b).toString(), "false")
            __check((a === c).toString(), "true")
        }
