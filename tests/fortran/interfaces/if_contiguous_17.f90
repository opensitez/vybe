! vybe-test: fortran/interfaces/if_contiguous_17
! origin: languages/fortran/tests/fortran/test_interfaces.rs
subroutine s(a)
real,contiguous::a(:)
end subroutine s
