using System;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Kompas6API5;
using Steps.NET;

namespace Project_Create_VN
{
	[ClassInterface(ClassInterfaceType.AutoDual)]
	public class Auto_create
	{
		// Имя библиотеки
		[return: MarshalAs(UnmanagedType.BStr)]
		public string GetLibraryName()
		{
			return "Auto__Create";
		}
		[return: MarshalAs(UnmanagedType.BStr)]
		public string ExternalMenuItem(short number, ref short itemType, ref short command)
		{
			string result = string.Empty;
			itemType = 1; // "MENUITEM"
			switch (number)
			{
				case 1:
					result = "Создание ВН + ЛУ + НДЗ";
					command = 1;
					break;

				case 2:
					result = "Продолжение следует...";
					command = 2;
					break;

				case 3:
					command = -1;
					itemType = 3; // "ENDMENU"
					break;
			}
			return result;
		}
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			KompasObject kompas = (KompasObject)kompas_;

			switch ((int)command)
			{
				case 1:
					Form1 _Form1 = new Form1();
					_Form1.ShowDialog();
					break;

				case 2:
					kompas.ksMessage("Нереализовано!");
					break;
			}
		}
		#region COM Registration
		[ComRegisterFunction]
		public static void RegisterKompasLib(Type t)
		{
			try
			{
				RegistryKey regKey = Registry.LocalMachine;
				string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
				regKey = regKey.OpenSubKey(keyName, true);
				regKey.CreateSubKey("Kompas_Library");
				regKey = regKey.OpenSubKey("InprocServer32", true);
				regKey.SetValue(null, System.Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\mscoree.dll");
				regKey.Close();
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("При регистрации класса для COM-Interop произошла ошибка:\n{0}", ex));
			}
		}

		// Эта функция удаляет раздел Kompas_Library из реестра
		[ComUnregisterFunction]
		public static void UnregisterKompasLib(Type t)
		{
			RegistryKey regKey = Registry.LocalMachine;
			string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
			RegistryKey subKey = regKey.OpenSubKey(keyName, true);
			subKey.DeleteSubKey("Kompas_Library");
			subKey.Close();
		}
		#endregion
	}
}
