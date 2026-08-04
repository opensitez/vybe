! vybe-test: fortran/stop_error_status_signaling/stop_error_status_signaling_error_stop_in_subroutine
! origin: languages/fortran/tests/fortran/test_stop_error_status_signaling.rs
program stop_error_status_signaling_error_stop_in_subroutine
call guard(.true.)
print *, 'after'
contains
            subroutine guard(flag)
                logical, intent(in) :: flag
                if (flag) error stop 2
            end subroutine guard
        end program stop_error_status_signaling_error_stop_in_subroutine
