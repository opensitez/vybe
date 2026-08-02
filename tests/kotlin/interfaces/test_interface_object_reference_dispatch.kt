// vybe-test: kotlin/interfaces/test_interface_object_reference_dispatch
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Status {
            fun code(): Int
        }

        class Offline : Status {
            override fun code(): Int = 0
        }

        class Online : Status {
            override fun code(): Int = 1
        }

        fun describe(status: Status): String {
            return if (status.code() == 0) "off" else "on"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe(Offline())).toString(), "off")
            __check((describe(Online())).toString(), "on")
        }
