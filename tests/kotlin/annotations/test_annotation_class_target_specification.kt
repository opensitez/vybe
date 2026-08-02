// vybe-test: kotlin/annotations/test_annotation_class_target_specification
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Target(AnnotationTarget.CLASS, AnnotationTarget.FUNCTION)
        @Retention(AnnotationRetention.RUNTIME)
        annotation class Role(val name: String)

        @Role("service")
        class Service {
            @Role("entry")
            fun start(): String = "ready"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val service = Service()
            __check((service.start()).toString(), "ready")
        }
