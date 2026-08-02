// vybe-test: kotlin/inheritance_dispatch/test_implementation_of_multiple_interfaces_is_dispatched_as_expected
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

interface Read {
            fun read(): String = "read"
        }

        interface Write {
            fun write(): String = "write"
        }

        class Device : Read, Write {
            override fun read(): String = "device-read"
            override fun write(): String = "device-write"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val device: Read = Device()
            val writer: Write = Device()
            __check((device.read()).toString(), "device-read")
            __check((writer.write()).toString(), "device-write")
        }
