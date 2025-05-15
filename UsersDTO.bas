Attribute VB_Name = "UsersDTO"


    ' Private fields (could use m prefix or keep p prefix as you prefer)
    Private mUserNumber As String
    Private mUsername As String
    Private mUserLevel As String
    Private mStatus As String
    Private mIsSessionActive As Boolean
    Private mRecordDate As Date

    ' Property Get and Let Methods for each property

    Public Property Get UserNumber() As String
        UserNumber = mUserNumber
    End Property

    Public Property Let UserNumber(ByVal value As String)
        mUserNumber = value
    End Property

    Public Property Get Username() As String
        Username = mUsername
    End Property

    Public Property Let Username(ByVal value As String)
        mUsername = value
    End Property

    Public Property Get UserLevel() As String
        UserLevel = mUserLevel
    End Property

    Public Property Let UserLevel(ByVal value As String)
        mUserLevel = value
    End Property

    Public Property Get Status() As String
        Status = mStatus
    End Property

    Public Property Let Status(ByVal value As String)
        mStatus = value
    End Property

    Public Property Get IsSessionActive() As Boolean
        IsSessionActive = mIsSessionActive
    End Property

    Public Property Let IsSessionActive(ByVal value As Boolean)
        mIsSessionActive = value
    End Property

    Public Property Get RecordDate() As Date
        RecordDate = mRecordDate
    End Property

    Public Property Let RecordDate(ByVal value As Date)
        mRecordDate = value
    End Property


