// vybe-test: kotlin/annotations/test_annotation_on_object_declaration
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Deprecated("legacy")
        class Notifier {
            companion object {
                fun ping(): String {
                    return "pong"
                }
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Notifier.ping()).toString(), "pong")
        }
