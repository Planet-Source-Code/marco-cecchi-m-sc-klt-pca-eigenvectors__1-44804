Attribute VB_Name = "KLT"
Option Explicit
Option Base 1

Const N = 2
Private Vector(N) As Double
Private A(N, N) As Double

'This is the VB source code for an eigenvector analysis routine.
'The routine tri_diag() is called first to convert a real symmetric matrix (such as a covariance matrix)
'to tri-diagonal form.
'Then calc_eigenstructure() is called to generate a matrix of eigenvectors, sorted in decending
'order of eigenvalue magnitude. When applied to an image covariance matrix, these routines can
'be used as the basis for the Hotelling Transform, which is the discrete form of the Karhunen-Loeve Expansion.
Private Sub main()
    Dim i As Long, j As Long
    
    Vector(1) = 1: Vector(2) = 1
    For i = 1 To N
        For j = 1 To i
            A(i, j) = Vector(i) * Vector(j)
        Next j
    Next i
A(1, 1) = 3
A(2, 1) = -5
A(1, 2) = -5
A(2, 2) = 3

    Dim StrDebug As String
    Debug.Print "Covariance Matrix"
    For i = 1 To N
        StrDebug = ""
        For j = 1 To N
            StrDebug = StrDebug + CStr(A(i, j)) + " "
        Next j
        Debug.Print StrDebug
    Next i
    
    Dim res As Boolean
    res = KLT(A)
    Debug.Print CStr(res)
    End
End Sub

Public Function KLT(ByRef SymmMat() As Double) As Boolean
    Dim tol As Double
    Dim macheps As Double
    Dim N As Long
    Dim i As Long, j As Long
    Dim d() As Double, e() As Double, A() As Double
    
    N = UBound(SymmMat(), 1)
    ReDim d(N)
    ReDim e(N)
    ReDim A(N, N)
    
    tol = 0.000001
    Call tri_diag(SymmMat, d, e, A, N, tol)
    Dim StrDebug As String
    Debug.Print "Tridiagonal Matrix"
    For i = 1 To N
        StrDebug = ""
        For j = 1 To N
            StrDebug = StrDebug + CStr(A(i, j)) + " "
        Next j
        Debug.Print StrDebug
    Next i
    
    macheps = 0.000001
    KLT = calc_eigenstructure(d, e, A, N, macheps)
    ' d eigenvalues vector
    Debug.Print "Eigenvalues vector"
    StrDebug = ""
    For i = 1 To N
            StrDebug = StrDebug + CStr(d(i)) + " "
    Next i
    Debug.Print StrDebug
    ' A eigenvectors matrix
    Debug.Print "Eigenvectors matrix"
    For i = 1 To N
        StrDebug = ""
        For j = 1 To N
            StrDebug = StrDebug + CStr(A(i, j)) + " "
        Next j
        Debug.Print StrDebug
    Next i
End Function

'**********************************************************************
'  routine to tri-diagonalize a real symmetric matrix
'        uses Householder 's method
'tri_diag( A, d, e, A, L, tol)
'A - an LxL sized matrix containing the symmetric matrix to be analyzed,
'      such as a covariance matrix (in C, allocate one longer, i.e., alloc length LxL+1 )
'd - a length L vector that passes results to the next routine
'     (in C, allocate one longer, i.e., alloc length L+1 ); needs no initialization
'e - a length L vector that passes results to the next routine
'     (in C, allocate one longer, i.e., alloc length L+1 ); needs no initialization
'z - an LxL matrix that holds the tri-diagonalized version of Cxd upon return;
'    needed to pass to the next routine (in C, allocate one longer, i.e., alloc length LxL+1 );
'    needs no initialization
'tol - tolerance for checking nearness to zero;
'     I found 1.0e-6 to be sufficient for my applications but you may need a smaller value.
'
'  NOTE: code was taken from an old ALGOL program and contains goto's
'**********************************************************************
Private Sub tri_diag(ByRef A() As Double, _
                        ByRef d() As Double, _
                        ByRef e() As Double, _
                        ByRef z() As Double, _
                        ByVal N As Long, _
                        ByVal tol As Double)
Dim i As Long, j As Long, k As Long, l As Long
Dim f As Double, g  As Double, h  As Double, hh  As Double

    i = 1
    While i <= N 'for( i = 1; i <= n; i++ )
        j = 1
        While j <= i 'for( j = 1; j <= i; j++ )
            z(i, j) = A(i, j)
            
            j = j + 1
        Wend
        
        i = i + 1
    Wend
    
    i = N
    While i >= 2 'for( i = n; i >= 2; i--)
        l = i - 2
        f = z(i, i - 1)

        g = 0#
        k = 1
        While k <= l 'for( k = 1; k <= l; k++ )
            g = g + z(i, k) * z(i, k)
            
            k = k + 1
        Wend
        h = g + f * f

        If g <= tol Then
            e(i) = f
            h = 0#
            GoTo skip
        End If

        l = l + 1
        If f >= 0# Then
            g = -Sqr(h)
            e(i) = -Sqr(h)
        Else
            g = Sqr(h)
            e(i) = Sqr(h)
        End If
        
        h = h - f * g
        z(i, i - 1) = f - g
        f = 0#
        
        j = 1
        While j <= l 'for( j = 1; j <= l; j++ )
            z(j, i) = z(i, j) / h
            g = 0#
            k = 1
            While k <= j 'for( k = 1; k <= j; k++)
                g = g + z(j, k) * z(i, k)
                
                k = k + 1
            Wend
            k = j + 1
            While k <= l 'for( k = j+1; k <= l; k++ )
                g = g + z(k, j) * z(i, k)
                
                k = k + 1
            Wend
            e(j) = g / h
            f = f + g * z(j, i)
            
            j = j + 1
        Wend
        
        hh = f / (h + h)
        j = 1
        While j <= l 'for( j = 1; j <= l; j++ )
            f = z(i, j)
            g = e(j) - hh * f
            e(j) = g
            k = 1
            While k <= j 'for( k = 1; k <= j; k++ )
                z(j, k) = z(j, k) - f * e(k) - g * z(i, k)
                
                k = k + 1
            Wend
            
            j = j + 1
        Wend
skip:
        d(i) = h
        
        i = i - 1
    Wend

    d(1) = 0#
    e(1) = 0#
    i = 1
    While i <= N 'for( i = 1; i <=n; i++ )
        l = i - 1
        If d(i) <> 0# Then
            j = 1
            While j <= l 'for( j = 1; j <= l; j++ )
                g = 0#
                k = 1
                While k <= l 'for( k = 1; k <= l; k++ )
                    g = g + z(i, k) * z(k, j)
                    
                    k = k + 1
                Wend
                k = 1
                While k <= l 'for( k = 1; k <= l; k++ )
                    z(k, j) = z(k, j) - g * z(k, i)
                    
                    k = k + 1
                Wend
                
                j = j + 1
            Wend
        End If
        d(i) = z(i, i)
        z(i, i) = 1#
        j = 1
        While j <= l 'for( j = 1; j <= l; j++ )
            z(i, j) = 0#
            z(j, i) = 0#
            
            j = j + 1
        Wend
        
        i = i + 1
    Wend
End Sub

'**********************************************************************
'    routine to find eigenstructure of real tri-diagonal matrix
'                 uses QL algorithm
'calc_eigenstructure( d, e, z, L, macheps )
'd - vector as above; on return it holds the eigenvalues in sorted order (smallest to largest)
'e - vector as above; used to pass info into this routine from the previous one
'z - matrix as above;
'   on return it holds the array of eigenvectors as columns in order
'   corresponding to the eigenvalues in d
'macheps - an iteration error tolerance parameter;
'          I found 1.0e-16 to work well in my applications; you may have to adjust
'          this if you have convergence problems.
'
'          returns  0: success      -1: failure to converge
'
'  NOTE: code was taken from an old ALGOL program and contains goto's
'**********************************************************************
Private Function calc_eigenstructure(ByRef d() As Double, _
                        ByRef e() As Double, _
                        ByRef z() As Double, _
                        ByVal N As Long, _
                        ByVal macheps As Double) As Boolean

Dim i As Long, j As Long, k As Long, l As Long, m As Long
Dim b As Double, c As Double, f As Double, g As Double, h As Double
Dim p As Double, r As Double, s As Double

    i = 2
    While i <= N 'for( i = 2; i <=n; i++ )
        e(i - 1) = e(i)
        i = i + 1
    Wend

    e(N) = 0#
    b = 0#
    f = 0#

    l = 1
    While l <= N 'for( l = 1; l <= n; l++ )
        j = 0
        h = macheps * (Abs(d(l)) + Abs(e(l)))
        If b < h Then b = h

        m = l
        While m <= N  'for( m = l; m <= n; m++ )
            If Abs(e(m)) <= b Then GoTo cont

            m = m + 1
        Wend

cont:
        If m = l Then GoTo root
nextit:
        If j = 30 Then
            calc_eigenstructure = False
            Exit Function
        End If

        j = j + 1
        p = (d(l + 1) - d(l)) / (2# * e(l))
        r = Sqr(p * p + 1)
        If p < 0# Then
            h = d(l) - e(l) / (p - r)
        Else
            h = d(l) - e(l) / (p + r)
        End If

        i = l
        While i <= N 'for( i = l; i <= n; i++ )
            d(i) = d(i) - h

            i = i + 1
        Wend

        f = f + h
        p = d(m)
        c = 1#
        s = 0#

        i = m - 1
        While i >= l 'for( i = m-1; i >= l; i-- )
            g = c * e(i)
            h = c * p

            If Abs(p) >= Abs(e(i)) Then
                c = e(i) / p
                r = Sqr(c * c + 1)
                e(i + 1) = s * p * r
                s = c / r
                c = 1# / r
            Else
                c = p / e(i)
                r = Sqr(c * c + 1)
                e(i + 1) = s * e(i) * r
                s = 1# / r
                c = c / r
            End If

            p = c * d(i) - s * g
            d(i + 1) = h + s * (c * g + s * d(i))

            k = 1
            While k <= N 'for( k = 1; k <= n; k++ )
                h = z(k, i + 1)
                z(k, i + 1) = s * z(k, i) + c * h
                z(k, i) = c * z(k, i) - s * h

                k = k + 1
            Wend

            i = i - 1
        Wend

        e(l) = s * p
        d(l) = c * p
        If (Abs(e(l)) > b) Then GoTo nextit
root:
        d(l) = d(l) + f

        l = l + 1
    Wend
    
    'order the eigenvectors
    i = 1
    While i <= N 'for( i = 1; i <= n; i++ )
        k = i
        p = d(i)

        j = i + 1
        While j <= N 'for( j = i+1; j <= n; j++ )
            If d(j) < p Then
                k = j
                p = d(j)
            End If

            j = j + 1
        Wend

        If k <> i Then
            d(k) = d(i)
            d(i) = p

            j = 1
            While j <= N 'for( j = 1; j <= n; j++ )
                p = z(j, i)
                z(j, i) = z(j, k)
                z(j, k) = p

                j = j + 1
            Wend
        End If

        i = i + 1
    Wend

    calc_eigenstructure = True
End Function
