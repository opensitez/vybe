// vybe-test: kotlin/random/test_random_default_generator_exists
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((kotlin.random.Random.Default != null).toString(), "true")
        }
