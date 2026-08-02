// vybe-test: kotlin/annotations/test_annotation_with_typealias
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Deprecated("old alias")
        class Greeting {
            val message: String = "hi"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val msg = Greeting()
            __check((msg.message).toString(), "hi")
        }
