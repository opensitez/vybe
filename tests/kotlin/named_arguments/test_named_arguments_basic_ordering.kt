// vybe-test: kotlin/named_arguments/test_named_arguments_basic_ordering
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun format(label: String, count: Int): String {
            return label + ":" + count
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((format(count = 2, label = "k")).toString(), "k:2")
        }
