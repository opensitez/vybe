// vybe-test: kotlin/kotlin_arrays_creation/test_array_of_any_casting_and_mutation
// origin: languages/kotlin/tests/kotlin/test_kotlin_arrays_creation.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val mix = arrayOf<Any>("x", 1, true)
            val tail = mix.drop(1)
            __check((tail.joinToString("|")).toString(), "1|true")
        }
