// vybe-test: kotlin/interfaces/test_interface_inheritance_property_override_chain
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Read {
            val source: String
        }

        interface Cache : Read {
            override val source: String
            fun open(): String = "cached:" + source
        }

        class SourceFile : Cache {
            override val source: String = "in-memory"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item: Cache = SourceFile()
            __check((item.source).toString(), "in-memory")
            __check((item.open()).toString(), "cached:in-memory")
        }
