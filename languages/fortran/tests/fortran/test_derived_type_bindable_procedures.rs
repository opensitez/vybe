use super::helpers::run_prints;

#[test]
fn test_derived_type_bindable_procedures_sets_state() {
    let out = run_prints(
        r#"
program test_derived_type_bindable_procedures
    type :: counter
        integer :: value = 1
    contains
        procedure :: inc => counter_inc
    end type

    type(counter) :: c
    call c%inc(4)
    print *, c%value

contains
    subroutine counter_inc(self, delta)
        class(counter), intent(inout) :: self
        integer, intent(in) :: delta
        self%value = self%value + delta
    end subroutine
end program test_derived_type_bindable_procedures
"#,
    );

    assert_eq!(out, vec!["5"]);
}
