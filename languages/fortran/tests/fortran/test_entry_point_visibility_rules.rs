use super::helpers::run_prints;

#[test]
fn test_entry_point_visibility_rules_calls_public_entry() {
    let out = run_prints(
        r#"
module entry_points
    public :: public_entry
    private :: private_entry

    contains

    integer function public_entry()
        public_entry = 9
    end function

    integer function private_entry()
        private_entry = 1
    end function
end module

program test_entry_point_visibility_rules
    use entry_points
    print *, public_entry()
end program test_entry_point_visibility_rules
"#,
    );

    assert_eq!(out, vec!["9"]);
}
