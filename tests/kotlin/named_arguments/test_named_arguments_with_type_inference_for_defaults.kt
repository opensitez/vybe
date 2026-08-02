// vybe-test: kotlin/named_arguments/test_named_arguments_with_type_inference_for_defaults
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun make(items: List<Int> = listOf(1, 2), label: String): String {
            return label + ":" + items.size
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((make(label = "x")).toString(), "x:2")
            __check((make(items = listOf(1), label = "y")).toString(), "y:1")
        }
