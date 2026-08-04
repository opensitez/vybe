! vybe-test: fortran/modules/dimension_array
! origin: languages/fortran/tests/fortran/test_modules.rs
program t
integer, dimension(5) :: arr
arr(1) = 10
print *, arr(1)
end program t
