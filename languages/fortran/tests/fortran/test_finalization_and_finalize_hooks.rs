use super::helpers::run_prints;

#[test]
fn test_finalization_and_finalize_hooks_cleanup_signal() {
    let out = run_prints(
        r#"
program test_finalization_and_finalize_hooks
    type :: scoped
        integer :: flag = 7
    contains
        final :: finalize_scoped
    end type

    block
        type(scoped) :: item
        print *, item%flag
    end block

contains
    subroutine finalize_scoped(self)
        type(scoped), intent(inout) :: self
        print *, self%flag
    end subroutine
end program test_finalization_and_finalize_hooks
"#,
    );

    assert_eq!(out, vec!["7", "7"]);
}
