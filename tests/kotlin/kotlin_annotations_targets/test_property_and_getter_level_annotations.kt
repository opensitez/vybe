// vybe-test: kotlin/kotlin_annotations_targets/test_property_and_getter_level_annotations
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotations_targets.rs

@Target(AnnotationTarget.PROPERTY, AnnotationTarget.FIELD)
        annotation class Visible

        class Box {
            @Visible
            var value: Int = 7
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Box()
            __check((b.value).toString(), "7")
        }
