// vybe-test: kotlin/data_classes/test_data_class_ordering_by_to_string_is_deterministic
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Pair(val a: Int, val b: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Pair(2, 1)
            val b = Pair(10, 3)
            val list = listOf(a, b)
            __check((list.sortedBy { it.a }.joinToString(";") { it.toString() }).toString(), "Pair(a=2, b=1);Pair(a=10, b=3)")
        }
