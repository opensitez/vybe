// vybe-test: kotlin/annotations/test_annotation_on_class_and_member
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Deprecated
        class Legacy {
            fun name(): String = "legacy"
        }

        @Suppress("UNUSED_PARAMETER")
        fun tagged(@Deprecated code: Int): String {
            return "tagged"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val legacy = Legacy()
            __check((legacy.name()).toString(), "legacy")
            __check((tagged(1)).toString(), "tagged")
        }
