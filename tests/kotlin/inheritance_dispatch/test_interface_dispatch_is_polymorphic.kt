// vybe-test: kotlin/inheritance_dispatch/test_interface_dispatch_is_polymorphic
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

interface Reader {
            fun read(): String
        }

        class A : Reader {
            override fun read(): String = "a"
        }

        class B : Reader {
            override fun read(): String = "b"
        }

        fun emit(readers: Array<Reader>): String {
            var total = ""
            for (reader in readers) {
                total += reader.read()
            }
            return total
        }

        fun main() {
            println(emit(arrayOf(A(), B())))
        }

