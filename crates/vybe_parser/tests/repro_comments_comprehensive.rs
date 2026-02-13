use vybe_parser::parse_program;

#[test]
fn test_end_sub_with_comment_variations() {
    let cases = vec![
        "Sub Foo()\nEnd Sub ' Comment",
        "Sub Foo()\nEnd Sub 'Comment",
        "Sub Foo()\nEnd Sub      ' Comment",
        "Sub Foo()\nEnd Sub ' Comment \n",
    ];

    for case in cases {
        let result = parse_program(case);
        assert!(result.is_ok(), "Failed to parse: '{}' -> {:?}", case, result.err());
    }
}
