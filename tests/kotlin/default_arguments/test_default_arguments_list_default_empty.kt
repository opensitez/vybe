// vybe-test: kotlin/default_arguments/test_default_arguments_list_default_empty
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun join(items: List<String> = listOf()): String = items.joinToString("-")
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("<" + join() + ">").toString(), "<>")
            __check((join(listOf("a", "b"))).toString(), "a-b")
        }
