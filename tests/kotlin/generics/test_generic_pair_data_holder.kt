// vybe-test: kotlin/generics/test_generic_pair_data_holder
// origin: languages/kotlin/tests/kotlin/test_generics.rs

class Holder<K, V>(private val key: K, private val value: V) {
            fun parts(): String {
                return key.toString() + ":" + value.toString()
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Holder("x", 7).parts()).toString(), "x:7")
            __check((Holder(2, true).parts()).toString(), "2:true")
        }
