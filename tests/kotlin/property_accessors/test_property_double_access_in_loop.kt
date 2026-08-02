// vybe-test: kotlin/property_accessors/test_property_double_access_in_loop
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Count {
            var n = 0
                private set
            fun inc() { n += 1 }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Count()
            c.inc()
c.inc()
            __check((c.n).toString(), "2")
        }
