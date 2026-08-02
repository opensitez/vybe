// vybe-test: kotlin/kotlin_destructuring_maps/test_destructure_component_functions_order
// origin: languages/kotlin/tests/kotlin/test_kotlin_destructuring_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val e = mapOf("z" to 9).entries.first()
            __check((e.component1()).toString(), "z")
            __check((e.component2()).toString(), "9")
        }
