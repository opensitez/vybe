! vybe-test: fortran/select_type_polymorphic_matching/typeof_integer_vector
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

! `typeof(x)` as a CALLABLE FUNCTION is not valid Fortran — not F2018 and
! not F2023, which adds TYPEOF only as a DECLARATION type-specifier
! (`typeof(x) :: y`). gfortran leaves `_typeof_` undefined at link. The type
! inquiry it was reaching for is spelled with an unlimited polymorphic and
! SELECT TYPE, which is real Fortran and actually exercises the machinery.

program t
    integer :: v(3) = [1, 2, 3]
    character(len=16) :: tname
    tname = type_name(v(1))
    if (trim(tname) /= "integer") then
        print *, "FAIL: want [integer] got [", trim(tname), "]"
        stop 1
    end if
contains
    function type_name(v) result(r)
        class(*), intent(in) :: v
        character(len=16) :: r
        select type (v)
        type is (integer)
            r = "integer"
        type is (real)
            r = "real"
        type is (logical)
            r = "logical"
        class default
            r = "other"
        end select
    end function type_name
end program t
