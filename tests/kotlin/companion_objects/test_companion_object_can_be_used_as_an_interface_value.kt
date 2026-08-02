// vybe-test: kotlin/companion_objects/test_companion_object_can_be_used_as_an_interface_value
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

interface Named {
            fun name(): String
        }

        class Factory {
            companion object : Named {
                override fun name(): String = "factory"
            }
        }

        fun label(source: Named): String = source.name()

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source: Named = Factory.Companion
            __check((label(source)).toString(), "factory")
            __check((label(Factory.Companion)).toString(), "factory")
        }
