###############################################################################
#
# Powershell Class to deal with Metex and Voltcraft M3850D Digital Multimeters
#
# created: 08.11.2024 Michael Albert info@michlstechblog.info
#
# License GPL V2
#
###############################################################################


# [System.IO.Ports.SerialPort]::GetPortNames()
# $COMPort=new-object System.IO.Ports.SerialPort
# $COMPort.BaudRate=1200
# $COMPort.DataBits=7
# $COMPort.PortName="COM5"
# $COMPort.Parity=[System.IO.Ports.Parity]::None
# $COMPort.Stopbits=[System.IO.Ports.StopBits]::Two
# $COMPort.Handshake=[System.IO.Ports.Handshake]::None
# 
# $iMaxTimeToReadIn100ms=2000
# $iMustBytesInQueue=14
# $COMPort.Open()
# 
# while($true){
#     $COMPort.Write("D")
#     do{
#         $iBytesInQueue=$COMPort.BytesToRead
#         $iMaxTimeToReadIn100ms--
#         write-host  "." -nonewline
#         Start-Sleep -Milliseconds 100
#     }
#     while($iMaxTimeToReadIn100ms -ge 0 -and $COMPort.BytesToRead -lt $iMustBytesInQueue)
#     write-host ("Bytes to read: {0}" -f $COMPort.BytesToRead)
#     $COMPort.ReadExisting()
# }
# 
# $COMPort.CLose()
# DC -000.0  mV
# DC  000.0  mV
# DC -000.1  mV
# DC -0.000   V
# FR  0.000 KHz
# LO   rdy
# LO - rdy
# OH  298.0KOhm
# OH   O.L MOhm
# OH  2.815MOhm
# OH  25.94MOhm
# OH  15.79MOhm
# OH  24.72MOhm
# OH  1.834MOhm
# OH  1.938MOhm
# DI  0000   mV
# DI  1740   mV
# DI  0663   mV
# DI  0742   mV
# DI  1532   mV
# DI  1934   mV
# CA  0.004  nF
# CA  0.005  nF
# HF  0010
# HF  0000
# TE -2969    C
# TE -  OL    C
# DC -494.8  uA
# DC -1.484  mA
# DC -000.1  uA
# DC -00.95  mA
# DC -00.01  mA
# DC -00.02  mA
# DC  00.28  mA
# DC -0.102   A
# DC -0.001   A
enum MeasuringUnit
{
	VoltageDC
	VoltageAC
	CurrentDC
	CurrentAC
	Frequency
	Resistance
	Capacity
	Temperature
	Diode
	Logic
	NotImplementedYet
	Undefined
}
enum MeasureResult
{
	OK
	TimeOut
	InvalidResponseLength
	COMPortNotReady
	COMPortNotOpen
	NotImplementedYet
	UnitRecognitionError
	ValueConvertingError
	Overload
}
class Measure
{
	[MeasuringUnit]$MeasuringUnit=[MeasuringUnit]::Undefined
	[single]$Value=0
	[string]$RawMeasureString=""
	[MeasureResult]$Result
	[string]$Unit=""
	[string]$RawValue=0
	[string]$RawUnit=""
	Measure()
	{
	}
}
class MetexM3850D
{

	hidden [string]$_COMPort="";
	hidden [int]$_BaudRate=1200
	hidden [int]$_DataBits=7
	hidden [System.IO.Ports.Parity]$_Parity=[System.IO.Ports.Parity]::None
	hidden [System.IO.Ports.StopBits]$_Stopbits=[System.IO.Ports.StopBits]::Two
	hidden [System.IO.Ports.Handshake]$_Handshake=[System.IO.Ports.Handshake]::None
	
	hidden [int]$MaxTimeToReadIn100ms=30
	hidden [int]$ReadTimeIn100ms=0
	hidden [int]$MustBytesInQueue=14

    hidden [Measure]$_Measure=$null
	
	[System.IO.Ports.SerialPort]$COMPort
	
	hidden [System.Text.RegularExpressions.Regex]$RegExSplitOffMeasureUnit=$null
	static [string]$SplitOffMeasureUnitregEx='(.*)(V|A|F|Ohm|Hz)$'

	MetexM3850D([string]$COMPort)
	{
		$this._COMPort=$COMPort
		$this._Measure=new-object Measure
		if(!([System.IO.Ports.SerialPort]::GetPortNames() -contains $this._COMPort))
		{
			write-warning ("Selected COM Port {0} does not exists! Available Ports: {1}" -f $this._COMPort,([String]::Join(",",[System.IO.Ports.SerialPort]::GetPortNames())))
		}
	}
	[System.IO.Ports.SerialPort] Open()
	{
		$this.RegExSplitOffMeasureUnit=New-Object System.Text.RegularExpressions.Regex([MetexM3850D]::SplitOffMeasureUnitregEx)
		try
		{
			# $this._Measure=new-object Measure
			$this.COMPort=new-object System.IO.Ports.SerialPort
			$this.COMPort.BaudRate=$this._BaudRate
			$this.COMPort.DataBits=$this._DataBits
			$this.COMPort.PortName=$this._COMPort
			$this.COMPort.Parity=$this._Parity
			$this.COMPort.Stopbits=$this._Stopbits
			$this.COMPort.Handshake=$this._Handshake			
			$this.COMPort.Open()
			return $this.COMPort 
		}	
		catch
		{
			if($this.COMPort -and $this.COMPort.IsOpen)
			{
				$this.COMPort.Close()
			}
			throw ("Metex library. Cannot open COM port {0}. Error: {1}" -f $this._COMPort,$_.Exception.Message)
		}
		return $null
	}
	[void] Close()
	{
		if($this.COMPort)
		{
			$this.COMPort.Close()
		}
	}
	[Measure] ReadMeasureSynchron()
	{
		if(!$this.COMPort)
		{
			
			$this._Measure.MeasuringUnit=[MeasuringUnit]::Undefined
			$this._Measure.Value=0
			$this._Measure.RawMeasureString="ERROR"
			$this._Measure.Result=[MeasureResult]::COMPortNotReady
			return $this._Measure
		}	
		if(!$this.COMPort.IsOpen)
		{
			$this._Measure.MeasuringUnit=[MeasuringUnit]::Undefined
			$this._Measure.Value=0
			$this._Measure.RawMeasureString="COM Port not open"
			$this._Measure.Result=[MeasureResult]::COMPortNotOpen
			return $this._Measure			
		}
		$this.COMPort.Write("D")
		$iBytesInQueue=0
		$this.ReadTimeIn100ms=$this.MaxTimeToReadIn100ms
		# Read Serial
		do{
		    $iBytesInQueue=$this.COMPort.BytesToRead
		    $this.ReadTimeIn100ms--
		    #write-host  "." -nonewline
		    Start-Sleep -Milliseconds 100
		}
		while($this.ReadTimeIn100ms -gt 0 -and $iBytesInQueue -lt $this.MustBytesInQueue)
		# Timeout occured?
		if($this.ReadTimeIn100ms -eq 0)
		{
			$this._Measure.MeasuringUnit=[MeasuringUnit]::Undefined
			$this._Measure.Value=0
			$this._Measure.RawMeasureString="Timeout occured!"
			$this._Measure.Result=[MeasureResult]::TimeOut	
			$this._Measure.RawValue=""
			$this.COMPort.DiscardInBuffer()			
		}
		# Byte count incorrect?
		elseif($iBytesInQueue -ne 14)
		{
			$this._Measure.MeasuringUnit=[MeasuringUnit]::Undefined
			$this._Measure.Value=0
			$this._Measure.RawMeasureString="Invalid character count: {0}!" -f $iBytesInQueue
			$this._Measure.Result=[MeasureResult]::InvalidResponseLength				
			$this._Measure.RawValue=""
			$this.COMPort.DiscardInBuffer()
		}
		#  Read OK
		else
		{
			# OK
			$this._Measure.RawMeasureString=$this.COMPort.ReadExisting()
			# Split off RAW String
			$MeasureType=$this._Measure.RawMeasureString.Substring(0,2)
			$UnitRaw=$this._Measure.RawMeasureString.Substring(9,4).Trim()
			$this._Measure.RawUnit=$UnitRaw
			$ValueRaw=$this._Measure.RawMeasureString.Substring(3,6).Trim()
			$this._Measure.RawValue=$ValueRaw
			# Split Unit and Exponent
			$oRegExSplitOffMeasureUnitMatch=$this.RegExSplitOffMeasureUnit.Match($UnitRaw)
			if(!$oRegExSplitOffMeasureUnitMatch.Success)
			{
				$this._Measure.MeasuringUnit=[MeasuringUnit]::Undefined
				$this._Measure.Value=0
				$this._Measure.RawMeasureString="Not implemented!"
				$this._Measure.Result=[MeasureResult]::UnitRecognitionError		
				return $this._Measure.Result
			}

			# Determine Exponent
			$sExponent=$oRegExSplitOffMeasureUnitMatch.Groups[1].Value
			$Mulitplier=1
			# Get Exponet from Unit String
			switch -casesensitive ($sExponent)
			{
				"m" {$Mulitplier=[Math]::Pow(10, -3)}
				"u" {$Mulitplier=[Math]::Pow(10, -6)}
				"n" {$Mulitplier=[Math]::Pow(10, -9)}
				"K" {$Mulitplier=[Math]::Pow(10, 3)}
				"M" {$Mulitplier=[Math]::Pow(10, 6)}
				"T" {$Mulitplier=[Math]::Pow(10, 9)}
			}
			# Unit
			$Unit=$oRegExSplitOffMeasureUnitMatch.Groups[2].Value
			$this._Measure.Unit=$Unit
			# Determine Measure type
			switch($MeasureType)
			{
				"DC" {
					if($Unit -eq "V")
					{
						$this._Measure.MeasuringUnit=[MeasuringUnit]::VoltageDC
					}
					elseif($Unit -eq "A")
					{
						$this._Measure.MeasuringUnit=[MeasuringUnit]::CurrentDC
					}
				}
				"AC" {
					if($Unit -eq "V")
					{
						$this._Measure.MeasuringUnit=[MeasuringUnit]::VoltageAC
					}
					elseif($Unit -eq "A")
					{
						$this._Measure.MeasuringUnit=[MeasuringUnit]::CurrentAC
					}					
				}
				"OH" {$this._Measure.MeasuringUnit=[MeasuringUnit]::Resistance}
				"CA" {$this._Measure.MeasuringUnit=[MeasuringUnit]::Capacity}
				"FR" {$this._Measure.MeasuringUnit=[MeasuringUnit]::Frequency}
				"DI" {$this._Measure.MeasuringUnit=[MeasuringUnit]::Diode}
				default {
					$this._Measure.MeasuringUnit=[MeasuringUnit]::NotImplementedYet
					$this._Measure.Value=0
					$this._Measure.RawMeasureString=("Measuretype {0} not implemented yet!" -f $MeasureType)
					$this._Measure.Result=[MeasureResult]::NotImplementedYet
					return $this._Measure.Result
				}
			}
			# Value
			try
			{
				# Special case Overload
				# write-host "RAW:" $ValueRaw
				if($ValueRaw -eq "OL" -or $ValueRaw -eq "O.L")
				{
					$this._Measure.Value=0
					$this._Measure.RawMeasureString="Measure Overload"
					$this._Measure.Result=[MeasureResult]::Overload
				}
				else
				{
					$this._Measure.Value=[System.Convert]::ToSingle($ValueRaw.Replace(".",","))
					$this._Measure.Result=[MeasureResult]::OK
				}	
			}
			catch
			{
				$this._Measure.MeasuringUnit=[MeasuringUnit]::Undefined
				$this._Measure.Value=0
				$this._Measure.RawMeasureString="Converting error! Raw value:" -f $ValueRaw
				$this._Measure.Result=[MeasureResult]::ValueConvertingError					
			}			
			# Apply Exponent to Value
			$this._Measure.Value*=$Mulitplier
		}
		return $this._Measure
	}
}
# Usage
# . .\libmetex.ps1  or . D:\temp\libmetex.ps1
# $Metex=new-object MetexM3850D("COM5")
# $SerialPort=$Metex.Open() | out-null
# $Metex.ReadMeasureSynchron()
# $Metex.Close()
