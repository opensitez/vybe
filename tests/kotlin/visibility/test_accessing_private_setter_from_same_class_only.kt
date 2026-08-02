// vybe-test: kotlin/visibility/test_accessing_private_setter_from_same_class_only
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

class Counter {
            var value: Int = 0
                private set

            fun setValue(next: Int) {
                value = next
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counter = Counter()
            counter.setValue(11)
            __check((counter.value).toString(), "11")
        }
