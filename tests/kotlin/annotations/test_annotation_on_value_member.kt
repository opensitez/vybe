// vybe-test: kotlin/annotations/test_annotation_on_value_member
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

class Counter {
            @Deprecated("counter")
            val total = 4
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((Counter().total).toString(), "4") }
