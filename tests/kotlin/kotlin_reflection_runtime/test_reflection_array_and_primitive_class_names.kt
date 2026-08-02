// vybe-test: kotlin/kotlin_reflection_runtime/test_reflection_array_and_primitive_class_names
// origin: languages/kotlin/tests/kotlin/test_kotlin_reflection_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Int::class.simpleName).toString(), "Int")
            __check((IntArray::class.simpleName).toString(), "IntArray")
            __check((Array<Int>::class.simpleName).toString(), "Array")
            __check((String::class.qualifiedName?.endsWith("kotlin.String")).toString(), "true")
        }
