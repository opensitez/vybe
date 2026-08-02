// vybe-test: kotlin/interfaces/test_interface_anonymous_object_with_capture
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Supplier {
            fun value(): String
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val prefix = "hello "
            val supplier = object : Supplier {
                override fun value(): String = prefix + "world"
            }
            __check((supplier.value()).toString(), "hello world")
        }
