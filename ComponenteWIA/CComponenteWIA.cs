
using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using WIA;

namespace ComponenteWIA
{
	#region estruturas
	public struct WIADeviceInfo
	{
		public string DeviceID;
		public string Name;
		public Dictionary<string, string> Propriedades;
		public WIADeviceInfo(string DeviceID, string Name)
		{
			this.DeviceID = DeviceID;
			this.Name = Name;
			this.Propriedades = new Dictionary<string, string>();
		}
	}

	// https://msdn.microsoft.com/en-us/library/windows/desktop/ms630313(v=vs.85).aspx
	public enum WIADeviceInfoProp
	{
		DeviceID = 2,
		Manufacturer = 3,
		Description = 4,
		Type = 5,
		Port = 6,
		Name = 7,
		Server = 8,
		RemoteDevID = 9,
		UIClassID = 10,
	}


	// http://www.papersizes.org/a-paper-sizes.htm
	public enum WIAPageSize
	{
		A4, // 8.3 x 11.7 in  (210 x 297 mm)
		Letter, // 8.5 x 11 in (216 x 279 mm)
		Legal, // 8.5 x 14 in (216 x 356 mm)
	}

	public enum WIAScanQuality
	{
		Preview,
		Final,
	}
	#endregion

	class WIAScanner
	{
		#region constants
		const string wiaFormatBMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}";
		const string WIA_DEVICE_PROPERTY_PAGES_ID = "3096";
		const string WIA_SCAN_BRIGHTNESS_PERCENTS = "6154";
		const string WIA_SCAN_CONTRAST_PERCENTS = "6155";
		const string WIA_SCAN_COLOR_MODE = "6146";
		const int DEVICE_PROPERTY_DOCUMENT_HANDLING_CAPABILITIES_ID = 3086;
		const int DEVICE_PROPERTY_DOCUMENT_HANDLING_STATUS_ID = 3087;
		const int DEVICE_PROPERTY_DOCUMENT_HANDLING_SELECT_ID = 3088;
		const int DEVICE_PROPERTY_PAGES_ID = 3096;

		public const int WIA_ERROR_GENERAL_ERROR = 1;
		public const int WIA_ERROR_PAPER_JAM = 2;
		public const int WIA_ERROR_PAPER_EMPTY = 3;
		public const int WIA_ERROR_PAPER_PROBLEM = 4;
		public const int WIA_ERROR_OFFLINE = 5;
		public const int WIA_ERROR_BUSY = 6;
		public const int WIA_ERROR_WARMING_UP = 7;
		public const int WIA_ERROR_USER_INTERVENTION = 8;
		public const int WIA_ERROR_ITEM_DELETED = 9;
		public const int WIA_ERROR_DEVICE_COMMUNICATION = 10;
		public const int WIA_ERROR_INVALID_COMMAND = 11;
		public const int WIA_ERROR_INCORRECT_HARDWARE_SETTING = 12;
		public const int WIA_ERROR_DEVICE_LOCKED = 13;
		public const int WIA_ERROR_EXCEPTION_IN_DRIVER = 14;
		public const int WIA_ERROR_INVALID_DRIVER_RESPONSE = 15;
		public const int WIA_S_NO_DEVICE_AVAILABLE = 21;

		#endregion

		#region internal_class
		class ScanSettings
		{
			public int dpi = 150;
			public int color = 2; //4 is black-white, gray is 2, color is 1
			public bool adf = true;
			public bool tryFlatbed = false; //try flatbed if adf fails
		}
		class WIA_DPS_DOCUMENT_HANDLING_SELECT
		{
			public const uint FEEDER = 0x00000001;
			public const uint FLATBED = 0x00000002;
			public const uint DUPLEX = 0x00000004;
			public const uint AUTO_ADVANCE = 0x00000200;
		}
		class WIA_DPS_DOCUMENT_HANDLING_STATUS
		{
			public const uint FEED_READY = 0x00000001;
		}
		class WIA_PROPERTIES
		{
			public const uint WIA_RESERVED_FOR_NEW_PROPS = 1024;
			public const uint WIA_DIP_FIRST = 2;
			public const uint WIA_DPA_FIRST = WIA_DIP_FIRST + WIA_RESERVED_FOR_NEW_PROPS;
			public const uint WIA_DPC_FIRST = WIA_DPA_FIRST + WIA_RESERVED_FOR_NEW_PROPS;
			//
			// Scanner only device properties (DPS)
			//
			public const uint WIA_DPS_FIRST = WIA_DPC_FIRST + WIA_RESERVED_FOR_NEW_PROPS;
			public const uint WIA_DPS_DOCUMENT_HANDLING_STATUS = WIA_DPS_FIRST + 13;
			public const uint WIA_DPS_DOCUMENT_HANDLING_SELECT = WIA_DPS_FIRST + 14;
			public const uint WIA_DPS_DOCUMENT_HANDLING_PICTURES_REMAINING = 2051;

		}
		#endregion

		#region public_members
		public static HashSet<string> Mensagens = new HashSet<string>();
		public static string CaminhoArquivo = Path.GetTempPath();
		public static Dictionary<string, ComponenteWIA.WIADeviceInfo> Dispositivos = new Dictionary<string, WIADeviceInfo>();
		public enum TipoLeituraDocumento : int
		{
			Feeder = 1,
			FlatBed = 2,
			FeederDuplex = 5
		}
		#endregion

		#region public_Methods

		/// <summary>
		/// Inicia a digitalização dos documentos depositados no dispositivo selecionado
		/// </summary>
		/// <param name="scannerId"></param>
		/// <param name="tamanho"></param>
		/// <param name="tipoDigitalizacao"></param>
		/// <returns></returns>
		public static List<KeyValuePair<string, Image>> Scan(string scannerId, WIAPageSize tamanho, TipoLeituraDocumento tipoDigitalizacao)
		{
			WIA.Device device = LocalizaDispositio(scannerId);

			String description = device.Properties["Name"].get_Value().ToString();

			if (description.ToLower().Contains("brother") || description.Contains("Canon MF4500"))
			{
				return DigitalizaItensEspecifico(scannerId, WIAScanQuality.Final, tamanho, tipoDigitalizacao);
			}
			else
			{
				return DigitalizaItens(scannerId, WIAScanQuality.Final, tamanho, tipoDigitalizacao);
			}
		}

		/// <summary>
		/// Retorna a lista de dispositivos disponíveis
		/// </summary>
		/// <returns></returns>
		public static IEnumerable<WIADeviceInfo> ListaDispositivos()
		{
			List<string> devices = new List<string>();
			WIA.DeviceManager manager = new WIA.DeviceManager();
			// https://msdn.microsoft.com/en-us/library/windows/desktop/ms630313(v=vs.85).aspx
			foreach (WIA.DeviceInfo info in manager.DeviceInfos)
			{
				if (info.Type == WIA.WiaDeviceType.ScannerDeviceType)
				{

					WIADeviceInfo mDevInfo = new WIADeviceInfo(info.DeviceID, info.Properties["Name"].get_Value().ToString());
					Dispositivos.Add(info.DeviceID, mDevInfo);
					yield return mDevInfo;
				}

			}
		}

		#endregion

		#region private_methods

		/// <summary>
		/// 
		/// </summary>
		/// <param name="scannerId"></param>
		/// <param name="qualidade"></param>
		/// <param name="tamanho"></param>
		/// <param name="tipo"></param>
		/// <returns></returns>
		private static List<KeyValuePair<string, Image>> DigitalizaItens(string scannerId, WIAScanQuality qualidade, WIAPageSize tamanho, TipoLeituraDocumento tipo)
		{
			List<KeyValuePair<string, Image>> images = new List<KeyValuePair<string, Image>>();
			//Dictionary<string,Image> images = new Dictionary<string,Image>();
			bool hasMorePages = true;
			WIA.Item item;

			Mensagens = new HashSet<string>();
			while (hasMorePages)
			{
				// select the correct scanner using the provided scannerId parameter
				WIA.Device device = LocalizaDispositio(scannerId);

				try
				{
					//SetWIAProperty(device.Properties, WIA_DEVICE_PROPERTY_PAGES_ID, 1);
					//Escolha da forma de scaneamento, Feeder ou Flatbad
					ConfiguraTipoScan(ref device, tipo);

					item = device.Items[1] as WIA.Item;

					AjustaPropriedadesDispositivo(item, 0, 0, 0, 0, 1, qualidade, tamanho);
					// scan image

					WIA.ICommonDialog wiaCommonDialog = new WIA.CommonDialog();
					WIA.ImageFile image = (WIA.ImageFile)wiaCommonDialog.ShowTransfer(item, WIA.FormatID.wiaFormatPNG, false);

					if (image != null)
					{
						// save to  file
						string fileName = SalvarPNG(image);
						images.Add(new KeyValuePair<string, Image>(fileName, Image.FromFile(fileName)));

					}

					//determine if there are any more pages waiting
					WIA.Property documentHandlingSelect = null;
					WIA.Property documentHandlingStatus = null;


					foreach (WIA.Property prop in device.Properties)
					{

						if (prop.PropertyID == WIA_PROPERTIES.WIA_DPS_DOCUMENT_HANDLING_SELECT)
							documentHandlingSelect = prop;
						if (prop.PropertyID == WIA_PROPERTIES.WIA_DPS_DOCUMENT_HANDLING_STATUS)
							documentHandlingStatus = prop;

					}
					// assume there are no more pages
					hasMorePages = false;
					// may not exist on flatbed scanner but required for feeder
					if (documentHandlingSelect != null)
					{
						// check for document feeder
						if ((Convert.ToUInt32(documentHandlingSelect.get_Value()) & WIA_DPS_DOCUMENT_HANDLING_SELECT.FEEDER) != 0)
						{
							hasMorePages = ((Convert.ToUInt32(documentHandlingStatus.get_Value()) & WIA_DPS_DOCUMENT_HANDLING_STATUS.FEED_READY) != 0);
						}
					}
				}
				catch (Exception exc)
				{
					int error = CodigoErroWIA(exc);
					Mensagens.Add(DescricaoErro(exc));
					if (error == WIA_ERROR_PAPER_EMPTY) hasMorePages = false;
					if (error == 0) throw exc;
				}
				finally
				{
					item = null;

				}
			}
			return images;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="scannerId"></param>
		/// <param name="qualidade"></param>
		/// <param name="tamanho"></param>
		/// <param name="tipo"></param>
		/// <returns></returns>
		private static List<KeyValuePair<string, Image>> DigitalizaItensEspecifico(string scannerId, WIAScanQuality qualidade, WIAPageSize tamanho, TipoLeituraDocumento tipo)
		{
			List<KeyValuePair<string, Image>> images = new List<KeyValuePair<string, Image>>();
			//Dictionary<string,Image> images = new Dictionary<string,Image>();
			bool hasMorePages = true;
			WIA.Item item;


			Mensagens = new HashSet<string>();

			// select the correct scanner using the provided scannerId parameter
			WIA.DeviceManager manager = new WIA.DeviceManager();
			WIA.Device device = null;
			foreach (WIA.DeviceInfo info in manager.DeviceInfos)
			{
				if (info.DeviceID == scannerId)
				{
					device = info.Connect();
					//AcquireNormal(device);
					Dispositivos[scannerId].Propriedades.Clear();
					CarregaPropriedades(device, Dispositivos[scannerId].Propriedades);
					break;
				}
			}
			// device was not found
			if (device == null)
			{
				// enumerate available devices
				string availableDevices = "";
				foreach (WIA.DeviceInfo info in manager.DeviceInfos)
				{
					availableDevices += info.DeviceID + "\n";
				}

				// show error with available devices
				Mensagens.Add("Não foi possível conectar-se ao o dispositivo especificado\n" + availableDevices);
			}

			ConfiguraTipoScan(ref device, tipo);

			item = device.Items[1] as WIA.Item;

			AjustaPropriedadesDispositivo(item, 0, 0, 0, 0, 1, qualidade, tamanho);
			WIA.ICommonDialog wiaCommonDialog = new WIA.CommonDialog();
			while (hasMorePages)
			{

				try
				{
					//Some scanner need WIA_DPS_PAGES to be set to 1, otherwise all pages are acquired but only one is returned as ImageFile
					GravaPropriedade(ref device, DEVICE_PROPERTY_PAGES_ID, 1);

					//Scan image
					WIA.ImageFile image = (WIA.ImageFile)wiaCommonDialog.ShowTransfer(item, WIA.FormatID.wiaFormatPNG, false);
					if (image != null)
					{
						// save to  file
						string fileName = SalvarPNG(image);
						images.Add(new KeyValuePair<string, Image>(fileName, Image.FromFile(fileName)));

					}

					//determine if there are any more pages waiting
					WIA.Property documentHandlingSelect = null;
					WIA.Property documentHandlingStatus = null;


					foreach (WIA.Property prop in device.Properties)
					{

						if (prop.PropertyID == WIA_PROPERTIES.WIA_DPS_DOCUMENT_HANDLING_SELECT)
							documentHandlingSelect = prop;
						if (prop.PropertyID == WIA_PROPERTIES.WIA_DPS_DOCUMENT_HANDLING_STATUS)
							documentHandlingStatus = prop;

					}
					// assume there are no more pages
					hasMorePages = false;
					// may not exist on flatbed scanner but required for feeder
					if (documentHandlingSelect != null)
					{
						// check for document feeder
						if ((Convert.ToUInt32(documentHandlingSelect.get_Value()) & WIA_DPS_DOCUMENT_HANDLING_SELECT.FEEDER) != 0)
						{
							hasMorePages = ((Convert.ToUInt32(documentHandlingStatus.get_Value()) & WIA_DPS_DOCUMENT_HANDLING_STATUS.FEED_READY) != 0);
						}
					}


				}
				catch (Exception exc)
				{
					int error = CodigoErroWIA(exc);
					Mensagens.Add(DescricaoErro(exc));
					if (error == WIA_ERROR_PAPER_EMPTY) hasMorePages = false;
					if (error == 0) throw exc;
				}
			}





			device = null;
			return images;
		}


		/// <summary>
		/// 
		/// </summary>
		/// <param name="dispositivo"></param>
		/// <param name="idPropriedade"></param>
		/// <param name="valor"></param>
		private static void GravaPropriedade(ref Device dispositivo, int idPropriedade, object valor)
		{
			Property property = PesquisaPropriedade(dispositivo.Properties, idPropriedade);
			if (property != null)
				property.set_Value(valor);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="dispositivo"></param>
		/// <param name="idPropriedade"></param>
		/// <returns></returns>
		private static object LePropriedade(ref Device dispositivo, int idPropriedade)
		{
			Property property = PesquisaPropriedade(dispositivo.Properties, idPropriedade);

			return property != null ? property.get_Value() : null;
		}
		/// <summary>
		/// 
		/// </summary>
		/// <param name="dispositivo"></param>
		/// <param name="tipo"></param>
		private static void ConfiguraTipoScan(ref Device dispositivo, TipoLeituraDocumento tipo)
		{
			//bool canDuplex = dispositivo.
			//DeviceSettings.DocumentHandlingCapabilities.HasFlag(DocumentHandlingCapabilities.Dup)
			int requested = (int)tipo;
			int supported = (int)LePropriedade(ref dispositivo,
							 DEVICE_PROPERTY_DOCUMENT_HANDLING_CAPABILITIES_ID);
			if (supported > WIA_DPS_DOCUMENT_HANDLING_SELECT.AUTO_ADVANCE && requested == (int)TipoLeituraDocumento.FeederDuplex) //Device doesn't support Feed/duplex
			{
				Mensagens.Add("Configuração de digitalização dupla face não é suportada pelo dispositivo");
				requested = (int)TipoLeituraDocumento.Feeder;
			}

			if ((requested & supported) != 0)
			{
				if ((requested & (int)TipoLeituraDocumento.Feeder) != 0)
				{
					GravaPropriedade(ref dispositivo, DEVICE_PROPERTY_PAGES_ID, 1);
				}
				GravaPropriedade(ref dispositivo, DEVICE_PROPERTY_DOCUMENT_HANDLING_SELECT_ID, requested);
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="propriedades"></param>
		/// <param name="idPropriedade"></param>
		/// <returns></returns>
		private static Property PesquisaPropriedade(WIA.Properties propriedades, int idPropriedade)
		{
			foreach (Property property in propriedades)
				if (property.PropertyID == idPropriedade)
					return property;
			return null;
		}
		/// <summary>
		/// 
		/// </summary>
		/// <param name="imagem"></param>
		private static string SalvarPNG(ImageFile imagem)
		{
			string fileName = CaminhoUnico();
			ImageProcess imgProcess = new ImageProcess();
			object convertFilter = "Convert";
			string convertFilterID = imgProcess.FilterInfos.get_Item(ref convertFilter).FilterID;
			imgProcess.Filters.Add(convertFilterID, 0);
			GravaPropriedadeWIA(imgProcess.Filters[imgProcess.Filters.Count].Properties, "FormatID", WIA.FormatID.wiaFormatPNG);
			imagem = imgProcess.Apply(imagem);
			imagem.SaveFile(fileName);
			return fileName;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <returns></returns>
		private static string CaminhoUnico()
		{
			string save_dir = CaminhoArquivo;
			string currentFName;
			string save_fullpath;
			int docno = 0;
			// DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss-fffffff")
			do
			{
				currentFName = String.Format("scan-{0}.PNG", docno);
				save_fullpath = Path.Combine(save_dir, currentFName);
				docno++;
			} while (File.Exists(save_fullpath));
			return save_fullpath;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="properties"></param>
		/// <param name="nomePropriedade"></param>
		/// <param name="valorPropriedade"></param>
		private static void GravaPropriedadeWIA(WIA.IProperties propriedades,
					 object nomePropriedade, object valorPropriedade)
		{
			try
			{

				WIA.Property prop = propriedades.get_Item(ref nomePropriedade);
				prop.set_Value(ref valorPropriedade);
			}
			catch (Exception mEx)
			{
				Mensagens.Add(string.Format("Não foi possível alterar a propriedade {0} Para {1}, description: {2}", nomePropriedade.ToString(), valorPropriedade.ToString(), DescricaoErro(mEx)));
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="item"></param>
		/// <param name="resolucaoDPI"></param>
		/// <param name="pixelInicialEsquerdo"></param>
		/// <param name="pixelInicialInicio"></param>
		/// <param name="larguraPixel"></param>
		/// <param name="alturaPixel"></param>
		/// <param name="percentualBrilho"></param>
		/// <param name="percentualContraste"></param>
		/// <param name="modoCor"></param>
		private static void AjustaPropriedadesDispositivo(IItem item, int pixelInicialEsquerdo, int pixelInicialInicio,
			int percentualBrilho, int percentualContraste, int modoCor, WIAScanQuality qualidade, WIAPageSize tamanho)
		{
			const string WIA_SCAN_COLOR_MODE = "6146";
			const string WIA_HORIZONTAL_SCAN_RESOLUTION_DPI = "6147";
			const string WIA_VERTICAL_SCAN_RESOLUTION_DPI = "6148";
			const string WIA_HORIZONTAL_SCAN_START_PIXEL = "6149";
			const string WIA_VERTICAL_SCAN_START_PIXEL = "6150";
			const string WIA_HORIZONTAL_SCAN_SIZE_PIXELS = "6151";
			const string WIA_VERTICAL_SCAN_SIZE_PIXELS = "6152";
			const string WIA_SCAN_BRIGHTNESS_PERCENTS = "6154";
			const string WIA_SCAN_CONTRAST_PERCENTS = "6155";
			// adjust the scan settings
			int resolucaoDPI;
			int larguraPixel;
			int alturaPixel;

			switch (qualidade)
			{
				case WIAScanQuality.Preview:
					resolucaoDPI = 100;
					break;
				case WIAScanQuality.Final:
					resolucaoDPI = 100;
					break;
				default:
					throw new Exception("Qualidade da imagem inválida: " + qualidade.ToString());
			}
			switch (tamanho)
			{
				case WIAPageSize.A4:
					larguraPixel = (int)(8.3f * resolucaoDPI);
					alturaPixel = (int)(11f * resolucaoDPI);
					break;
				case WIAPageSize.Letter:
					larguraPixel = (int)(8.5f * resolucaoDPI);
					alturaPixel = (int)(11f * resolucaoDPI);
					break;
				case WIAPageSize.Legal:
					larguraPixel = (int)(8.5f * resolucaoDPI);
					alturaPixel = (int)(11f * resolucaoDPI);
					break;
				default:
					throw new Exception("Tamanho da página inválido: " + tamanho.ToString());
			}

			GravaPropriedadeWIA(item.Properties, WIA_HORIZONTAL_SCAN_RESOLUTION_DPI, resolucaoDPI);
			GravaPropriedadeWIA(item.Properties, WIA_VERTICAL_SCAN_RESOLUTION_DPI, resolucaoDPI);
			GravaPropriedadeWIA(item.Properties, WIA_HORIZONTAL_SCAN_START_PIXEL, pixelInicialEsquerdo);
			GravaPropriedadeWIA(item.Properties, WIA_VERTICAL_SCAN_START_PIXEL, pixelInicialInicio);
			GravaPropriedadeWIA(item.Properties, WIA_HORIZONTAL_SCAN_SIZE_PIXELS, larguraPixel);
			GravaPropriedadeWIA(item.Properties, WIA_VERTICAL_SCAN_SIZE_PIXELS, alturaPixel);
			GravaPropriedadeWIA(item.Properties, WIA_SCAN_BRIGHTNESS_PERCENTS, percentualBrilho);
			GravaPropriedadeWIA(item.Properties, WIA_SCAN_CONTRAST_PERCENTS, percentualContraste);
			GravaPropriedadeWIA(item.Properties, WIA_SCAN_COLOR_MODE, modoCor);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cx"></param>
		/// <returns></returns>
		private static int CodigoErroWIA(Exception cx)
		{
			if (cx.GetType() == typeof(System.Runtime.InteropServices.COMException))
			{
				System.Runtime.InteropServices.COMException mCx = (System.Runtime.InteropServices.COMException)cx;
				int origErrorMsg = mCx.ErrorCode;

				int errorCode = origErrorMsg & 0xFFFF;

				int errorFacility = ((origErrorMsg) >> 16) & 0x1fff;

				if (errorFacility == 33) //WIAFacility
				{
					return errorCode;
				}
				else
				{
					return -1;
				}
			}

			return 0;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="errorCode"></param>
		/// <returns></returns>
		private static string DescricaoErroPorCodigo(int errorCode)
		{
			string desc = null;

			switch (errorCode)
			{
				case (WIA_ERROR_GENERAL_ERROR):
					{
						desc = "Erro genérico ao conectar ao dispositivo";
						break;
					}
				case (WIA_ERROR_PAPER_JAM):
					{
						desc = "Verifique se o papel está preso no dispositivo";
						break;
					}
				case (WIA_ERROR_PAPER_EMPTY):
					{
						desc = "Não foram encontrados outros documentos para digitalização";
						break;
					}
				case (WIA_ERROR_PAPER_PROBLEM):
					{
						desc = "Há problemas com o papel do dispositivo";
						break;
					}
				case (WIA_ERROR_OFFLINE):
					{
						System.Threading.Thread.Sleep(2000);
						desc = "O dispositivo está desligado";
						break;
					}
				case (WIA_ERROR_BUSY):
					{
						System.Threading.Thread.Sleep(2000);
						desc = "O dispositivo está ocupado";
						break;
					}
				case (WIA_ERROR_WARMING_UP):
					{
						desc = "O dispositivo está aquecendo";
						break;
					}
				case (WIA_ERROR_USER_INTERVENTION):
					{
						desc = "O dispositivo requer uma intervenção manual";
						break;
					}
				case (WIA_ERROR_ITEM_DELETED):
					{
						desc = "O dispositivo não está mais disponível";
						break;
					}
				case (WIA_ERROR_DEVICE_COMMUNICATION):
					{
						desc = "Ocorreu um erro ao tentar se comunicar com o dispositivo";
						break;
					}
				case (WIA_ERROR_INVALID_COMMAND):
					{
						desc = "O dispositivo não reconhece o comando";
						break;
					}
				case (WIA_ERROR_INCORRECT_HARDWARE_SETTING):
					{
						desc = "O dispositivo possui uma configuração incorreta";
						break;
					}
				case (WIA_ERROR_DEVICE_LOCKED):
					{
						desc = "O dispositivo está em uso por outra aplicação";
						break;
					}
				case (WIA_ERROR_EXCEPTION_IN_DRIVER):
					{
						desc = "O driver do dispositivo retornou um erro";
						break;
					}
				case (WIA_ERROR_INVALID_DRIVER_RESPONSE):
					{
						desc = "O Driver do dispositivo retornou uma resposta inválida";
						break;
					}
				default:
					{
						desc = "Ocorreu um erro desconhecido ao conectar-se ao dispositivo";
						break;
					}
			}

			return desc;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="pEx"></param>
		/// <returns></returns>
		private static string DescricaoErro(Exception pEx)
		{
			string mRet = pEx.Message;
			int errorCode = CodigoErroWIA(pEx);
			if (errorCode != 0)
			{
				mRet = DescricaoErroPorCodigo(errorCode);
			}
			return mRet;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="device"></param>
		/// <returns></returns>
		private static void CarregaPropriedades(WIA.Device device, Dictionary<string, string> mList)
		{

			try
			{

				// check if a device was selected
				if (device != null)
				{
					// Print camera properties
					foreach (Property prop in device.Properties)
					{
						mList.Add(prop.Name + "(dispositivo)", "(Valor) = " + prop.get_Value().ToString() + "  (PropertyId = " + prop.PropertyID + ": IsReadOnly = " + prop.IsReadOnly + ")");
						// Update UI
					}

					// Print item properties
					foreach (Property prop in device.Items[1].Properties)
					{
						mList.Add(prop.Name + "(item 1- dispositivo)", "(PropertyID) = " + prop.PropertyID.ToString() + "  (IsReadOnly = " + prop.IsReadOnly.ToString() + ") \n");
					}

					// Print commands
					foreach (DeviceCommand com in device.Commands)
					{
						mList.Add(com.Name + "(commandos)", "(Descrição) = " + com.Description + "  (CommandId = " + com.CommandID + ")");
					}

					// Print events
					foreach (DeviceEvent evt in device.Events)
					{
						mList.Add(evt.Name + "(events)", "(Descrição) = " + evt.Description + "  (Type = " + evt.Type + ") \n");
					}

				}
			}
			catch (Exception ex)
			{
				Mensagens.Add(ex.Message + "WIA Error!");
			}

		}

		private static WIA.Device LocalizaDispositio(string scannerId)
		{
			WIA.DeviceManager manager = new WIA.DeviceManager();
			WIA.Device device = null;
			foreach (WIA.DeviceInfo info in manager.DeviceInfos)
			{
				if (info.DeviceID == scannerId)
				{
					device = info.Connect();
					//AcquireNormal(device);
					Dispositivos[scannerId].Propriedades.Clear();
					CarregaPropriedades(device, Dispositivos[scannerId].Propriedades);
					break;
				}
			}
			// device was not found
			if (device == null)
			{
				// enumerate available devices
				string availableDevices = "";
				foreach (WIA.DeviceInfo info in manager.DeviceInfos)
				{
					availableDevices += info.DeviceID + "\n";
				}

				// show error with available devices
				Mensagens.Add("Não foi possível conectar-se ao o dispositivo especificado\n" + availableDevices);
			}
			return device;
		}
		#endregion

	}
}
