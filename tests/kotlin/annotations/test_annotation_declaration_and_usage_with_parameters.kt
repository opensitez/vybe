// vybe-test: kotlin/annotations/test_annotation_declaration_and_usage_with_parameters
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

annotation class Tag(val label: String)

        @Tag("service")
        fun service() {
            __check(("service").toString(), "service")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            service()
        }
