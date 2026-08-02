// vybe-test: kotlin/kotlin_annotation_usage/test_annotation_with_enum_and_boolean_args
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotation_usage.rs

enum class Level { OFF, WARN, ON }

        @Target(AnnotationTarget.CLASS)
        annotation class Flag(val level: Level, val active: Boolean = true)

        @Flag(level = Level.ON, active = true)
        class Processor

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Processor::class.simpleName).toString(), "Processor")
        }
