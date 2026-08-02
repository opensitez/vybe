// vybe-test: kotlin/annotations/test_annotation_companion_function
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

class Factory {
            companion object {
                @Deprecated("legacy") fun create(): Int = 21
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((Factory.create()).toString(), "21") }
