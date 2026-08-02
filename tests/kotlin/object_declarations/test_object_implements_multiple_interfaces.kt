// vybe-test: kotlin/object_declarations/test_object_implements_multiple_interfaces
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

interface Named {
            fun name(): String
        }

        interface Versioned {
            fun version(): Int
        }

        object Metadata : Named, Versioned {
            override fun name(): String = "meta"
            override fun version(): Int = 1
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Metadata.name()).toString(), "meta")
            __check((Metadata.version()).toString(), "1")
        }
