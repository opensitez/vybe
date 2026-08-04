! vybe-test: fortran/interfaces/if_pointer_proc_arg_40
! origin: languages/fortran/tests/fortran/test_interfaces.rs
subroutine apply(f)
procedure() :: f
call f()
end subroutine apply
