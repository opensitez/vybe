// vybe-test: kotlin/kotlin_annotations_targets/test_multiple_compatible_annotations_on_same_target
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotations_targets.rs

@Target(AnnotationTarget.CLASS)
        annotation class A(val value: String)

        @Target(AnnotationTarget.CLASS)
        annotation class B(val value: Int)

        @A("layer")
        @B(3)
        class Tagged

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Tagged::class.simpleName).toString(), "Tagged")
        }
