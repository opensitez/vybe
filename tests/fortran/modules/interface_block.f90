! vybe-test: fortran/modules/interface_block
! origin: languages/fortran/tests/fortran/test_modules.rs
program t
interface
function add(a, b) result(r)
integer, intent(in) :: a, b
integer :: r
end function add
end interface
print *, "ok"
end program t
