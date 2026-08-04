! vybe-test: fortran/oop/oop_class_result_08
! origin: languages/fortran/tests/fortran/test_oop.rs
function f() result(r)
class(*), allocatable :: r
allocate(integer :: r)
end function f
