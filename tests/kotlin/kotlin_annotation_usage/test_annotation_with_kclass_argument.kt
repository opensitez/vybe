// vybe-test: kotlin/kotlin_annotation_usage/test_annotation_with_kclass_argument
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotation_usage.rs

@Target(AnnotationTarget.CLASS)
        annotation class DelegateType(val impl: kotlin.reflect.KClass<*>)

        @DelegateType(impl = String::class)
        class Storage

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Storage::class.simpleName).toString(), "Storage")
        }
