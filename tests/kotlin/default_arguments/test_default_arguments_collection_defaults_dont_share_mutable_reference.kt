// vybe-test: kotlin/default_arguments/test_default_arguments_collection_defaults_dont_share_mutable_reference
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun make(items: MutableList<Int> = mutableListOf(1, 2)): String {
            items.add(3)
            return items.joinToString(":")
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = mutableListOf(9)
            __check((make(a)).toString(), "9:3")
            __check((make()).toString(), "1:2:3")
        }
