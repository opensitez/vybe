// vybe-test: kotlin/object_declarations/test_object_without_state_is_reusable_singleton_reference
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

object Marker

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Marker === Marker).toString(), "true")
        }
