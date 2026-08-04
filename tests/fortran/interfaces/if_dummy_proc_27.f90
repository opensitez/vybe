! vybe-test: fortran/interfaces/if_dummy_proc_27
! origin: languages/fortran/tests/fortran/test_interfaces.rs
subroutine apply(f)
external f
call f()
end subroutine apply
