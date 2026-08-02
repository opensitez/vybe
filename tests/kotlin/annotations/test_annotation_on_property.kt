// vybe-test: kotlin/annotations/test_annotation_on_property
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

class Versioned {
            @Deprecated("legacy field")
            val tag = "v1"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v = Versioned()
            __check((v.tag).toString(), "v1")
        }
