// vybe-test: kotlin/property_accessors/test_property_lazy_reused
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Holder {
            var initCount = 0
            val data: Int by lazy {
                initCount += 1
                9
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = Holder()
            __check((h.initCount).toString(), "0")
            __check((h.data).toString(), "9")
            __check((h.data).toString(), "9")
            __check((h.initCount).toString(), "1")
        }
