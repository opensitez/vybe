// vybe-test: kotlin/annotations/test_annotation_meta_targets
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Target(AnnotationTarget.CLASS)
        @Retention(AnnotationRetention.RUNTIME)
        annotation class TypeTag

        @TypeTag
        class Marked

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item: Any = Marked()
            __check((item::class.simpleName).toString(), "Marked")
        }
