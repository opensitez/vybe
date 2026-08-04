! vybe-test: fortran/finalization_and_finalize_hooks/test_finalization_and_finalize_hooks_cleanup_signal
! origin: languages/fortran/tests/fortran/test_finalization_and_finalize_hooks.rs

program test_finalization_and_finalize_hooks
    type :: scoped
        integer :: flag = 7
    contains
        final :: finalize_scoped
    end type

    block
        type(scoped) :: item
        if ((item%flag) /= 7) then
    print *, "FAIL: want [7] got [", item%flag, "]"
    stop 1
end if
    end block

contains
    subroutine finalize_scoped(self)
        type(scoped), intent(inout) :: self
        if ((self%flag) /= 7) then
    print *, "FAIL: want [7] got [", self%flag, "]"
    stop 1
end if
    end subroutine
end program test_finalization_and_finalize_hooks
