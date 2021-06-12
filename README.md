AudioEndPointController
=======================

Adapted to be used as a DLL instead of an EXE as originally intended.

How to use in C#

	private const string EPC = "EndPointController.dll";
	[DllImport(EPC)] [return: MarshalAs(UnmanagedType.BStr)] extern static string GetList();
	[DllImport(EPC, CallingConvention = CallingConvention.Cdecl)] static extern IntPtr SetAudio(int deviceid);

	[...]

	Console.WriteLine(GetList());
	Console.WriteLine(SetAudio(4));