use crate::helpers::run_prints;

#[test]
fn test_index_of_substring_from_offset() {
    let out = run_prints(
        r#"
        fun main() {
            val text = "banana"
            println(text.indexOf("na"))
            println(text.indexOf("na", 3))
            println(text.lastIndexOf("na"))
        }
    "#,
    );
    assert_eq!(out, &["2", "4", "4"]);
}

#[test]
fn test_index_of_any_character_set() {
    let out = run_prints(
        r#"
        fun main() {
            val text = "apple,banana,carrot"
            println(text.indexOfAny(charArrayOf(',', 'a'), 0))
            println(text.indexOfAny(charArrayOf('-', '/')))
        }
    "#,
    );
    assert_eq!(out, &["0", "-1"]);
}

#[test]
fn test_last_index_of_any_character_set() {
    let out = run_prints(
        r#"
        fun main() {
            val text = "abXcdY"
            println(text.lastIndexOfAny(charArrayOf('X', 'Y', 'Z')))
            println(text.lastIndexOfAny(charArrayOf('Q')))
        }
    "#,
    );
    assert_eq!(out, &["0", "-1"]);
}

#[test]
fn test_starts_with_end_with() {
    let out = run_prints(
        r#"
        fun main() {
            val text = "KotlinLang"
            println(text.startsWith("Kot"))
            println(text.startsWith("tin", 2))
            println(text.endsWith("Lang"))
            println(text.endsWith("lang", ignoreCase = true))
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "true", "true"]);
}

#[test]
fn test_contains_with_ignore_case() {
    let out = run_prints(
        r#"
        fun main() {
            val text = "Kotlin"
            println(text.contains("lin"))
            println(text.contains("LIN", ignoreCase = true))
            println(text.contains('K'))
            println(text.contains('z'))
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "true", "false"]);
}

#[test]
fn test_substring_before_after_variants() {
    let out = run_prints(
        r#"
        fun main() {
            val text = "a=b=c"
            println(text.substringAfter("="))
            println(text.substringAfter("z", "none"))
            println(text.substringBefore("="))
            println(text.substringBeforeLast("="))
            println(text.substringAfterLast("="))
        }
    "#,
    );
    assert_eq!(out, &["b=c", "none", "a", "a=b", "c"]);
}

#[test]
fn test_substring_after_last_with_multiple_delimiters() {
    let out = run_prints(
        r#"
        fun main() {
            val text = "/home/user/docs/readme.txt"
            println(text.substringAfterLast("/"))
            println(text.substringBeforeLast("/"))
            println("abc".substringAfterLast("/", "na"))
        }
    "#,
    );
    assert_eq!(out, &["readme.txt", "/home/user/docs", "na"]);
}

#[test]
fn test_find_character_and_indices() {
    let out = run_prints(
        r#"
        fun main() {
            val text = "kotlin"
            println(text.indexOfFirst { it in 'a'..'z' })
            println(text.indexOfLast { it == 't' })
            println(text.find { it == 'i' })
            println(text.findLast { it == 'o' } ?: "none")
        }
    "#,
    );
    assert_eq!(out, &["0", "2", "i", "o"]);
}

#[test]
fn test_replace_and_replace_first_boundaries() {
    let out = run_prints(
        r#"
        fun main() {
            val text = "banana"
            println(text.replaceFirst("ba", "pa"))
            println(text.replace("na", "xx"))
            println(text.replace("na", "xx", false))
        }
    "#,
    );
    assert_eq!(out, &["panana", "baxxxx", "baxxxx"]);
}

#[test]
fn test_line_and_trimmed_queries() {
    let out = run_prints(
        r#"
        fun main() {
            val block = "\n a \n b\n"
            val lines = block.lines()
            val trimmed = block.trim()
            println(lines.size)
            println(trimmed)
            println(trimmed.isNotEmpty())
        }
    "#,
    );
    assert_eq!(out, &["4", "a \n b", "true"]);
}

#[test]
fn test_chunked_and_drop_last_search() {
    let out = run_prints(
        r#"
        fun main() {
            val source = "abcdef"
            println(source.chunked(2).joinToString("|"))
            println(source.takeLast(3))
            println(source.dropLast(2))
        }
    "#,
    );
    assert_eq!(out, &["ab|cd|ef", "def", "abcd"]);
}

#[test]
fn test_string_in_operator_uses_index() {
    let out = run_prints(
        r#"
        fun main() {
            val text = "kotlin"
            println('t' in text)
            println('x' in text)
            println(3 in 1..5)
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "true"]);
}
