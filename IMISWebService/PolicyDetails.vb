'Copyright (c) 2016-%CurrentYear% Swiss Agency for Development and Cooperation (SDC)
'
'The program users must agree to the following terms:
'
'Copyright notices
'This program is free software: you can redistribute it and/or modify it under the terms of the GNU AGPL v3 License as published by the 
'Free Software Foundation, version 3 of the License.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of 
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU AGPL v3 License for more details www.gnu.org.
'
'Disclaimer of Warranty
'There is no warranty for the program, to the extent permitted by applicable law; except when otherwise stated in writing the copyright 
'holders and/or other parties provide the program "as is" without warranty of any kind, either expressed or implied, including, but not 
'limited to, the implied warranties of merchantability and fitness for a particular purpose. The entire risk as to the quality and 
'performance of the program is with you. Should the program prove defective, you assume the cost of all necessary servicing, repair or correction.
'
'Limitation of Liability 
'In no event unless required by applicable law or agreed to in writing will any copyright holder, or any other party who modifies and/or 
'conveys the program as permitted above, be liable to you for damages, including any general, special, incidental or consequential damages 
'arising out of the use or inability to use the program (including but not limited to loss of data or data being rendered inaccurate or losses 
'sustained by you or third parties or a failure of the program to operate with any other programs), even if such holder or other party has been 
'advised of the possibility of such damages.
'
'In case of dispute arising out or in relation to the use of the program, it is subject to the public law of Switzerland. The place of jurisdiction is Berne.


Public Class PolicyDetails

    Private _ProductCode As String
    Private _ProductName As String
    Private _ExpiryDate As String
    Private _Status As String
    Private _DedType As Nullable(Of Double)
    Private _Ded1 As Nullable(Of Double)
    Private _Ded2 As Nullable(Of Double)
    Private _Ceiling1 As Nullable(Of Double)
    Private _Ceiling2 As Nullable(Of Double)
    Private _TotalAdmissionsLeft As Nullable(Of Integer)
    Private _TotalVisitsLeft As Nullable(Of Integer)
    Private _TotalConsultationsLeft As Nullable(Of Integer)
    Private _TotalSurgeriesLeft As Nullable(Of Integer)
    Private _TotalDelivieriesLeft As Nullable(Of Integer)
    Private _TotalAntenatalLeft As Nullable(Of Integer)
    Private _ConsultationAmountLeft As Nullable(Of Integer)
    Private _SurgeryAmountLeft As Nullable(Of Integer)
    Private _HospitalizationAmountLeft As Nullable(Of Integer)
    Private _AntenatalAmountLeft As Nullable(Of Integer)



    Public Property ProductName() As String
        Get
            ProductName = _ProductName
        End Get
        Set(ByVal value As String)
            _ProductName = value
        End Set
    End Property

    Public Property ExpiryDate() As String
        Get
            ExpiryDate = _ExpiryDate
        End Get
        Set(ByVal value As String)
            _ExpiryDate = value
        End Set
    End Property

    Public Property Status() As String
        Get
            Status = _Status
        End Get
        Set(ByVal value As String)
            _Status = value
        End Set
    End Property

    Public Property DedType() As Nullable(Of Double)
        Get
            DedType = _DedType
        End Get
        Set(ByVal value As Nullable(Of Double))
            _DedType = value
        End Set
    End Property

    Public Property Ded1() As Nullable(Of Double)
        Get
            Ded1 = _Ded1
        End Get
        Set(ByVal value As Nullable(Of Double))
            _Ded1 = value
        End Set
    End Property

    Public Property Ded2() As Nullable(Of Double)
        Get
            Ded2 = _Ded2
        End Get
        Set(ByVal value As Nullable(Of Double))
            _Ded2 = value
        End Set
    End Property

    Public Property Ceiling1() As Double
        Get
            Ceiling1 = _Ceiling1
        End Get
        Set(ByVal value As Double)
            _Ceiling1 = value
        End Set
    End Property

    Public Property Ceiling2() As Nullable(Of Double)
        Get
            Ceiling2 = _Ceiling2
        End Get
        Set(ByVal value As Nullable(Of Double))
            _Ceiling2 = value
        End Set
    End Property

    Public Property ProductCode() As String
        Get
            ProductCode = _ProductCode
        End Get
        Set(ByVal value As String)
            _ProductCode = value
        End Set
    End Property

    Public Property TotalAdmissionsLeft() As Nullable(Of Integer)
        Get
            Return _TotalAdmissionsLeft
        End Get
        Set(ByVal value As Nullable(Of Integer))
            _TotalAdmissionsLeft = value
        End Set
    End Property

    Public Property TotalVisitsLeft() As Nullable(Of Integer)
        Get
            Return _TotalVisitsLeft
        End Get
        Set(ByVal value As Nullable(Of Integer))
            _TotalVisitsLeft = value
        End Set
    End Property

    Public Property TotalConsultationsLeft() As Nullable(Of Integer)
        Get
            Return _TotalConsultationsLeft
        End Get
        Set(ByVal value As Nullable(Of Integer))
            _TotalConsultationsLeft = value
        End Set
    End Property

    Public Property TotalSurgeriesLeft() As Nullable(Of Integer)
        Get
            Return _TotalSurgeriesLeft
        End Get
        Set(ByVal value As Nullable(Of Integer))
            _TotalSurgeriesLeft = value
        End Set
    End Property

    Public Property TotalDelivieriesLeft() As Nullable(Of Integer)
        Get
            Return _TotalDelivieriesLeft
        End Get
        Set(ByVal value As Nullable(Of Integer))
            _TotalDelivieriesLeft = value
        End Set
    End Property

    Public Property TotalAntenatalLeft() As Nullable(Of Integer)
        Get
            Return _TotalAntenatalLeft
        End Get
        Set(ByVal value As Nullable(Of Integer))
            _TotalAntenatalLeft = value
        End Set
    End Property

    Public Property ConsultationAmountLeft() As Nullable(Of Integer)
        Get
            Return _ConsultationAmountLeft
        End Get
        Set(ByVal value As Nullable(Of Integer))
            _ConsultationAmountLeft = value
        End Set
    End Property

    Public Property SurgeryAmountLeft() As Nullable(Of Integer)
        Get
            Return _SurgeryAmountLeft
        End Get
        Set(ByVal value As Nullable(Of Integer))
            _SurgeryAmountLeft = value
        End Set
    End Property

    Public Property HospitalizationAmountLeft() As Nullable(Of Integer)
        Get
            Return _HospitalizationAmountLeft
        End Get
        Set(ByVal value As Nullable(Of Integer))
            _HospitalizationAmountLeft = value
        End Set
    End Property

    Public Property AntenatalAmountLeft() As Nullable(Of Integer)
        Get
            Return _AntenatalAmountLeft
        End Get
        Set(ByVal value As Nullable(Of Integer))
            _AntenatalAmountLeft = value
        End Set
    End Property
End Class

