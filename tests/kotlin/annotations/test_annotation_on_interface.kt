// vybe-test: kotlin/annotations/test_annotation_on_interface
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Deprecated("old interface")
        interface Marker {
            fun label(): String
        }

        class Tag : Marker {
            override fun label(): String {
                return "tagged"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m: Marker = Tag()
            __check((m.label()).toString(), "tagged")
        }
