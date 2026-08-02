// vybe-test: kotlin/kotlin_reflection_runtime/test_reflection_java_class_name
// origin: languages/kotlin/tests/kotlin/test_kotlin_reflection_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val cls = ProbeImpl::class.java
            __check((cls.name).toString(), "languages.kotlin.tests.kotlin.test_kotlin_reflection_runtime.ProbeImpl")
            __check((cls.canonicalName).toString(), "languages.kotlin.tests.kotlin.test_kotlin_reflection_runtime.ProbeImpl")
            __check((cls.simpleName).toString(), "ProbeImpl")
        }
