// vybe-test: kotlin/interfaces/test_interface_array_dispatch_across_types
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Token {
            fun kind(): String
        }

        class Alpha : Token {
            override fun kind(): String = "alpha"
        }

        class Beta : Token {
            override fun kind(): String = "beta"
        }

        fun main() {
            val tokens: Array<Token> = arrayOf(Alpha(), Beta(), Alpha())
            var alpha = 0
            var beta = 0
            for (token in tokens) {
                when (token.kind()) {
                    "alpha" -> alpha += 1
                    "beta" -> beta += 1
                }
            }
            println(alpha)
            println(beta)
        }

