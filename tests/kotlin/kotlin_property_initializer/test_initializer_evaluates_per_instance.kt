// vybe-test: kotlin/kotlin_property_initializer/test_initializer_evaluates_per_instance
// origin: languages/kotlin/tests/kotlin/test_kotlin_property_initializer.rs

var marker = 0

        fun step(): Int {
            marker = marker + 10
            return marker
        }

        class Token {
            val value = step()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = Token()
            val second = Token()
            __check((first.value).toString(), "10")
            __check((second.value).toString(), "20")
        }
