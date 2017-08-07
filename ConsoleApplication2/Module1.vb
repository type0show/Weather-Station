Imports MetasysSystemSecureDataAccess
Imports MSSOAPLib30
Imports MSXML2
Imports System.Threading.Thread
Imports System.Net
Imports System.IO

'  
'  Author: Alex Motl, Franklin Chen
'  Date: 8/30/2016, updated 7/5/2017
'  Company: UW-Madison Digital Controls Group
'  Data Provided by: Space Science & Engineering Center UW-Madison 
'

Module Module1

    Public JCISecurity As MSSDAAPI
    Public SoapClient As MSSOAPLib30.SoapClient30
    Public TimeSoapClient As MSSOAPLib30.SoapClient30
    Public LoginTime As Date
    Public LoginDeviceTime As Date

    'Converts HPa to inHg
    Public Const PressureConversionFactor As Double = 0.0295299
    'Converts m/s to mph
    Public Const VelocityConversionFactor As Double = 2.23693629
    'Converts mm to inches
    Public Const LengthConversionFactor As Double = 0.03937008

    Sub Main()

        Dim result As Long
        Dim retStatus As String
        Dim siteIP As String
        Dim userName As String
        Dim password As String
        Dim WsdlUrl As String

        'site ip and user login information
        siteIP = "10.84.48.194"
        userName = "MotherNature"
        password = "Mete0r0!0gy"

        'If the location or name of the weather points, change the following
        Dim weatherFolderRef As String = "Weather-NAE:Weather-NAE/Programming.Weather Data."
        Dim temperatureName As String = "DRY-BULB"
        Dim wetBulbTemperatureName As String = "WET-BULB"
        Dim dewPointName As String = "DEW-PT"
        Dim windDirectionName As String = "WIND-DIR"
        Dim windSpeedName As String = "WIND-SPD"
        Dim pressureName As String = "PRESSURE"
        Dim solarFluxName As String = "SOLAR-FLUX"
        Dim precipitationName As String = "PRECIPITATION"
        Dim relHumidityName As String = "HUMIDITY"
        Dim alarmName As String = "WEBLINK-ALM"
        Dim altimeterName As String = "ALTIMETER"
        Dim hourName As String = "HOUR"
        Dim minuteName As String = "MINUTE"

        'value is delay in seconds
        'NOTE: should be >5 to ensure each update is unique since AOSS updates website every 5 seconds
        'NOTE: In case we are sending too many requests to the AOSS website, increase the updateDelay 
        Dim updateDelay As Int32 = 60

        JCISecurity = New MSSDAAPI
        result = JCISecurity.LoginUser(siteIP, userName, password, retStatus)
        If result = 0 Then

            'initialize the SOAP client and TimeSoapClient
            SoapClient = New MSSOAPLib30.SoapClient30
            WsdlUrl = "http://" + siteIP + "/MetasysIII/WS/BasicServices.asmx?wsdl"
            SoapClient.MSSoapInit(WsdlUrl)
            SoapClient.HeaderHandler = JCISecurity.headerHandler

            Dim timeWsdlUrl As String
            TimeSoapClient = New MSSOAPLib30.SoapClient30
            timeWsdlUrl = "http://" + siteIP + "/MetasysIII/WS/TimeManagement/TimeService.asmx?wsdl"
            TimeSoapClient.MSSoapInit(timeWsdlUrl)

            While True
                Dim weatherLine As String
                Try
                    'returns the unparsed line of weather data from AOSS website
                    'NOTE: If the website was unreachable the "Catch - End Try" block will execute reading/writing the backup data from Chem-NAE1
                    'NOTE: If successful connection "Catch - End Try" block is skipped
                    weatherLine = ReadCurrentAOSSWeatherData()
                Catch ex As System.Net.WebException

                    Console.WriteLine("Unable to Connect: " + Now())
                    Wait(updateDelay)
                    Continue While

                End Try

                Console.WriteLine("Connected: " + Now())

                'parse string of current data from AOSS site
                Dim weatherData As String() = Split(weatherLine, ",")

                'populate weather data using the tokenized weather data string 
                'NOTE: the value of the ConversionFactor variables are at the top of the file
                '      along with what units they are converting 
                Dim temperature As Double = CelciusToFahrenheitConversion(weatherData(2))
                Dim hour As Double = ParseHour(weatherData(1))
                Dim minute As Double = ParseMinute(weatherData(1))
                Dim dewPoint As Double = CelciusToFahrenheitConversion(weatherData(4))
                Dim windDirection As Double = weatherData(6)
                Dim windSpeed As Double = VelocityConversionFactor * weatherData(5)
                Dim pressure As Double = PressureConversionFactor * weatherData(8)
                Dim solarFlux As Double = weatherData(10)
                Dim precipitation As Double = LengthConversionFactor * weatherData(7)
                Dim relativeHumidity As Double = weatherData(3)
                Dim altimeter As Double = weatherData(9)
                Dim wetBulbTemperature As Double = CalculateWetBulbTemp(weatherData(2), weatherData(4), weatherData(8))

                'Write the populated weather data to the Metasys points in Weather-NAE and also the BD alarm point to normal.
                WritePresentValue(weatherFolderRef + temperatureName, temperature)
                WritePresentValue(weatherFolderRef + dewPointName, dewPoint)
                WritePresentValue(weatherFolderRef + wetBulbTemperatureName, wetBulbTemperature)
                WritePresentValue(weatherFolderRef + windDirectionName, windDirection)
                WritePresentValue(weatherFolderRef + windSpeedName, windSpeed)
                WritePresentValue(weatherFolderRef + pressureName, pressure)
                WritePresentValue(weatherFolderRef + solarFluxName, solarFlux)
                WritePresentValue(weatherFolderRef + precipitationName, precipitation)
                WritePresentValue(weatherFolderRef + relHumidityName, relativeHumidity)
                WritePresentValue(weatherFolderRef + altimeterName, altimeter)
                WritePresentValue(weatherFolderRef + alarmName, 0.0)
                WritePresentValue(weatherFolderRef + hourName, hour)
                WritePresentValue(weatherFolderRef + minuteName, minute)


                'delay next update until specified delay
                'NOTE: AOSS updates their website every ~5 seconds
                'NOTE: this update delay is not the guarenteed update interval, since the establishing the connection to 
                '      the website also takes a variable amount time (~5-7sec)
                Wait(updateDelay)
            End While
        Else
            Console.WriteLine("Error: Failed to Log in User")
        End If

    End Sub

    Sub Wait(seconds As Int32)
        Sleep(seconds * 1000) ' argument of Sleep function is milliseconds, so 1000 = 1 second
    End Sub

    Sub WritePresentValue(itemReference As String, newValue As Double)

        Dim status As Long '
        Dim itemPriority As String
        Dim retStatus As String
        Dim currentTime As String

        currentTime = GetCurrentDeviceTime()
        status = JCISecurity.InitMethodAuthentication(currentTime, "WriteProperty", itemReference, retStatus)
        On Error Resume Next
        status = SoapClient.WriteProperty(itemReference, "Present Value", newValue, itemPriority, retStatus)

    End Sub

    Function ReadPresentValue(itemReference As String) As Double

        Dim stringValue As String
        Dim readValue As Double
        Dim reliability As String
        Dim itemPriority As String
        Dim status As Long
        Dim retstatus As String
        Dim currentTime As String

        currentTime = GetCurrentDeviceTime()
        status = JCISecurity.InitMethodAuthentication(currentTime, "ReadProperty", itemReference, retstatus)
        On Error Resume Next
        status = SoapClient.ReadProperty(itemReference, "Present Value", stringValue, readValue, reliability, itemPriority)
        Return readValue

    End Function

    Function GetCurrentDeviceTime() As String

        Dim currentTimeResult As IXMLDOMNodeList
        Dim currentTime As String
        currentTimeResult = TimeSoapClient.GetCurrentTime()
        currentTime = currentTimeResult.item(1).text
        Return currentTime

    End Function

    Function ReadCurrentAOSSWeatherData() As String

        Dim address As String = "http://metobs.ssec.wisc.edu/app/rig/tower/data/ascii?symbols=t:rh:td:spd:dir:accum_precip:p:altm:flux&begin=-00:05:00"

        Dim client As WebClient = New WebClient()
        Dim reader As StreamReader = New StreamReader(client.OpenRead(address))
        Dim cLine, pLine As String
        Do
            pLine = cLine
            cLine = reader.ReadLine()
        Loop Until cLine Is Nothing
        reader.Close()
        Return pLine

    End Function

    ' Author: Franklin Chen
    ' Date: 7/5/2017
    ' Reference: https://www.youtube.com/watch?v=rWS2mUZfd1s

    ' This function uses a iterative method that approximates the wet-bulb temperatures
    Function CalculateWetBulbTemp(Temp As Double, Td As Double, pressure As Double) As Double

        Dim Tw_Td_avg, es_Tw_Td_avg, Delta, Twet, wetBulbTemperature_f As Double
        Dim loop_index As Integer
        Dim PsychrometricConstant = 0.000665 * pressure / 10

        Twet = 0

        For loop_index = 1 To 10 Step 1
            Tw_Td_avg = (Td + Twet) / 2
            es_Tw_Td_avg = 0.611 * (Math.E) ^ (17.502 * Tw_Td_avg / (Tw_Td_avg + 240.97))
            Delta = 17.502 * 240.97 * es_Tw_Td_avg / ((240.97 + Tw_Td_avg) ^ 2)
            Twet = (PsychrometricConstant * Temp + Delta * Td) / (Delta + PsychrometricConstant)
        Next

        ' Temperature return in degrees Fahrenheit
        wetBulbTemperature_f = CelciusToFahrenheitConversion(Twet)
        Return wetBulbTemperature_f
    End Function

    ' Author: Franklin Chen
    ' Date: 7/26/2017
    Function ParseHour(time As String) As Double

        Dim timeComponents As String() = Split(time, ":")
        Dim hour As Double = timeComponents(0) - 5

        Return hour

    End Function
    ' Author: Franklin Chen
    ' Date: 7/26/2017
    Function ParseMinute(time As String) As Double

        Dim timeComponents As String() = Split(time, ":")
        Dim minute As Double = timeComponents(1)

        Return minute

    End Function

    Function CelciusToFahrenheitConversion(temp As Double) As Double
        Return ((9.0 / 5) * temp) + 32
    End Function

End Module
