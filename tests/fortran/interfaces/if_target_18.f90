! vybe-test: fortran/interfaces/if_target_18
! origin: languages/fortran/tests/fortran/test_interfaces.rs
subroutine s(x)
integer,target::x
end subroutine s
