// vybe-test: kotlin/equality_hashcode/test_list_indexof_uses_structural_equality
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Entry(val value: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(Entry(1), Entry(2))
            __check((list.indexOf(Entry(1))).toString(), "0")
            __check((list.indexOf(Entry(3))).toString(), "-1")
        }
