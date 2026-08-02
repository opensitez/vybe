// vybe-test: kotlin/properties/test_property_setter_validates_negative_values
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Score {
            private var raw = 0
            var value: Int
                get() = raw
                set(next) { raw = if (next < 0) 0 else next }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val score = Score()
            score.value = -4
            __check((score.value).toString(), "0")
            score.value = 7
            __check((score.value).toString(), "7")
        }
