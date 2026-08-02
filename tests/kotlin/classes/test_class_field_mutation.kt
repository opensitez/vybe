// vybe-test: kotlin/classes/test_class_field_mutation
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Counter(var count: Int) {
            fun inc() {
                count += 1
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Counter(10)
            c.inc()
            __check((c.count).toString(), "11")
        }
