// vybe-test: kotlin/type_casts/test_safe_cast_between_independent_interfaces
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

interface First { fun first(): Int }
        interface Second { fun second(): String }

        class Impl : First, Second {
            override fun first(): Int = 7
            override fun second(): String = "ok"
        }

        fun main() {
            val value: Any = Impl()
            val first = value as First
            val second = value as? Second
            if (second != null) {
                println(first.first().toString() + ":" + second.second())
            } else {
                println("missing")
            }
        }

