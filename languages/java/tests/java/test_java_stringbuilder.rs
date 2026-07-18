use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(stringbuilder_initial, "StringBuilder sb = new StringBuilder(); System.out.println(sb.length());", "0");
jt!(stringbuilder_append_int, "StringBuilder sb = new StringBuilder(); sb.append(12); System.out.println(sb.toString());", "12");
jt!(stringbuilder_append_text, "StringBuilder sb = new StringBuilder(); sb.append(\"a\"); sb.append(\"b\"); System.out.println(sb.toString());", "ab");
jt!(stringbuilder_chain, "StringBuilder sb = new StringBuilder().append(\"x\").append(\"y\"); System.out.println(sb);", "xy");
jt!(stringbuilder_capacity_greater_or_equal, "StringBuilder sb = new StringBuilder(); System.out.println(sb.capacity() >= 16);", "true");
jt!(stringbuilder_append_char, "StringBuilder sb = new StringBuilder(); sb.append('J'); System.out.println(sb.toString());", "J");
jt!(stringbuilder_insert, "StringBuilder sb = new StringBuilder(\"ace\"); sb.insert(1, \"b\"); System.out.println(sb.toString());", "abc");
jt!(stringbuilder_delete, "StringBuilder sb = new StringBuilder(\"abcd\"); sb.delete(1, 3); System.out.println(sb.toString());", "ad");
jt!(stringbuilder_replace, "StringBuilder sb = new StringBuilder(\"abcd\"); sb.replace(1, 3, \"ZZ\"); System.out.println(sb.toString());", "aZZd");
jt!(stringbuilder_reverse, "StringBuilder sb = new StringBuilder(\"abc\"); sb.reverse(); System.out.println(sb.toString());", "cba");
jt!(stringbuilder_set_length_shorter, "StringBuilder sb = new StringBuilder(\"abcde\"); sb.setLength(3); System.out.println(sb.toString());", "abc");
jt!(stringbuilder_set_length_longer, "StringBuilder sb = new StringBuilder(\"ab\"); sb.setLength(4); System.out.println(sb.length());", "4");
jt!(stringbuilder_char_at, "StringBuilder sb = new StringBuilder(\"zxy\"); System.out.println(sb.charAt(1));", "x");
jt!(stringbuilder_append_boolean, "StringBuilder sb = new StringBuilder(); sb.append(true); System.out.println(sb.toString());", "true");
jt!(stringbuilder_append_float, "StringBuilder sb = new StringBuilder(); sb.append(1.5); System.out.println(sb.toString());", "1.5");
jt!(stringbuilder_append_object, "StringBuilder sb = new StringBuilder(); Object o = null; sb.append(o); System.out.println(sb.toString());", "null");
jt!(stringbuilder_append_code_point, "StringBuilder sb = new StringBuilder(); sb.appendCodePoint(0x41); System.out.println(sb.toString());", "A");
jt!(stringbuilder_ensure_capacity_noop, "StringBuilder sb = new StringBuilder(); sb.ensureCapacity(50); System.out.println(sb.capacity() >= 50);", "true");
jt!(stringbuilder_sub_sequence, "StringBuilder sb = new StringBuilder(\"abcdef\"); System.out.println(sb.subSequence(1, 4).toString());", "bcd");
jt!(stringbuilder_delete_char_at, "StringBuilder sb = new StringBuilder(\"hello\"); sb.deleteCharAt(1); System.out.println(sb.toString());", "hllo");
jt!(stringbuilder_append_multiple_types, "StringBuilder sb = new StringBuilder(); sb.append(1).append(\"-\").append(2.0).append(true); System.out.println(sb.toString());", "1-2.0true");
jt!(stringbuilder_replace_slice, "StringBuilder sb = new StringBuilder(\"banana\"); sb.replace(1, 4, \"OO\"); System.out.println(sb.toString());", "bOOna");
jt!(stringbuilder_insert_at_end, "StringBuilder sb = new StringBuilder(\"end\"); sb.insert(sb.length(), \"!\"); System.out.println(sb.toString());", "end!");
jt!(stringbuilder_clear_like, "StringBuilder sb = new StringBuilder(\"abc\"); sb.setLength(0); System.out.println(sb.length());", "0");
jt!(stringbuilder_to_string_idempotent, "StringBuilder sb = new StringBuilder(\"x\"); String s = sb.toString(); System.out.println(s);", "x");
jt!(stringbuilder_append_line, "StringBuilder sb = new StringBuilder(); sb.append(\"a\"); sb.append(\"b\"); System.out.println(sb.length());", "2");
jt!(stringbuilder_append_char_array, "char[] chars = {'J','a','v','a'}; StringBuilder sb = new StringBuilder(); sb.append(chars); System.out.println(sb.toString());", "Java");
