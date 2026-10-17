//! Keep reference-output regression diagnosis independent of parser selection.

#[test]
fn output_parameter_programs_preserve_complete_peg_ast_and_spans() {
    for source in [
        "<?php parse_str('', $result); return is_array($result) ? 'array' : 'missing';",
        r#"<?php
            function parse_wrapper($input, &$output) { parse_str($input, $output); }
            $result = null;
            parse_wrapper('', $result);
            return is_array($result) ? 'array' : 'missing';
        "#,
        r#"<?php
            function parse_wrapper($input, &$output) { parse_str($input, $output); }
            function options($input, $defaults) {
                parse_wrapper($input, $parsed);
                return array_merge($defaults, $parsed);
            }
            $result = options('', ['optimize' => false]);
            return $result['optimize'] ? 'bad' : 'ok';
        "#,
    ] {
        let direct = vybe_language_php::parse(source).unwrap();
        let peg = vybe_language_php::parse_with_peg(source).unwrap();
        assert_eq!(format!("{direct:?}"), format!("{peg:?}"), "{source}");
    }
}
