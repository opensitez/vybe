// vybe-test: kotlin/class_delegation/test_iterator_forwarded_from_list_delegation
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

class ReadOnlyList(delegate: List<String>) : Iterable<String> by delegate

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val it: String = ReadOnlyList(listOf("a", "b")).joinToString("")
            __check((it).toString(), "ab")
        }
