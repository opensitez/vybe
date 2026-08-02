// vybe-test: kotlin/nullability/test_nullable_array_of_objects
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

class Holder(val value: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first: Holder? = Holder(5)
            val second: Holder? = null
            __check((first?.value ?: 0).toString(), "5")
            __check((second?.value ?: 9).toString(), "9")
        }
