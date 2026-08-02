// vybe-test: kotlin/type_inference/test_type_inference_with_reified_like_simulation
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

inline fun <reified T> typeName(value: T): String = value!!::class.simpleName ?: ""
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((typeName(1)).toString(), "Int")
            __check((typeName("x")).toString(), "String")
        }
