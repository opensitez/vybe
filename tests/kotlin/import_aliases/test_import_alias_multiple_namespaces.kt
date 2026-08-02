// vybe-test: kotlin/import_aliases/test_import_alias_multiple_namespaces
// origin: languages/kotlin/tests/kotlin/test_import_aliases.rs

import java.lang.StringBuilder as KotlinBuilder
        import java.util.StringTokenizer as Tokenizer

        fun main() {
            val builder = KotlinBuilder()
            builder.append("a").append("b")
            println(builder.toString())
            val tokenizer = Tokenizer("x,y", ",")
            var tokens = 0
            while (tokenizer.hasMoreTokens()) {
                tokenizer.nextToken()
                tokens += 1
            }
            println(tokens)
        }

