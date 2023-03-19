using System;
using System.Collections.Generic;
using System.Linq;
using Multicad;
using Multicad.CustomObjectBase;
using Multicad.DatabaseServices;
using Multicad.DatabaseServices.StandardObjects;
using Multicad.Geometry;
using Multicad.Mc3D;
using Multicad.Objects;
using Multicad.Runtime;
using Multicad.DataServices;

namespace DotNetSample
{
	/// <summary>
	/// Примеры работы с параметрическим объектом БД.
	/// Examples of working with parametric DB objects.
	/// </summary>
	[ContainsCommands]
	public class ParametricObjectSamples
	{
		/// <summary>
		/// Пример интерактивной вставки параметрического объекта в чертеж с предварительным выбором его типа.
		/// The sample command selecting DB object and its following interactively placing into active document.
		/// </summary>
		[CommandMethod("smpl_PlaceDbObject", CommandFlags.NoCheck | CommandFlags.NoPrefix)]
		public static void PlaceParamObj()
		{
			Connection DbCon = new Connection();
			ICollection<Element> DbEls = DbCon.ShowElementBrowseDialog(new ElementFilter());
			if(DbEls != null && DbEls.Count > 0)
			{
				McParametricObject parObj = new McParametricObject();
				parObj.Initialize(DbEls.FirstOrDefault().ID);
				parObj.PlaceObject();
			}
		}

		/// <summary>
		/// Пример вставки в чертеж параметрического объекта БД определенного типа и исполнения.
		/// The sample command placing specific DB parametric object into acrive document.
		/// </summary>
		[CommandMethod("smpl_PlaceDbObject2", CommandFlags.NoCheck | CommandFlags.NoPrefix)]
		public static void PlaceSpecParamObj()
		{
			// Задаём тип объекта 2D/3D, в нашем случае - 3D
			// Define the object's type, either 2D or 3D, 3D in our case
			bool b3D = true;

			McParametricObject parObj = new McParametricObject(b3D);
			parObj.DbEntity.AddToCurrentDocument();

			// Установка связи параметрического объекта с объектом БД
			// Init connection between the parametric object and the object in the DB. 
			// "Детали крепления/Общее машиностроение/Болты/С шестигранной головкой/ГОСТ 7795-70 Исп. 1, 3, 4"
			parObj.Initialize(0x426CA8520B06C000L);

			// Установка 3-го исполнения
			// Setup the 3'rd implementation
			parObj.SetImplementationAndProcess("Implementation 3");

			// Установить диаметр болта = М16 и длины = 100мм
			// Setup bolt's diameter to M16 and the length = 100mm
			List<ExValue> ParamsToChange = new List<ExValue>() { new ExValue("dr", 16), new ExValue("L", 100) };
			parObj.Change(ParamsToChange, true);

			// Поместить объект в нужную точку
			// Place the object into desired position
			Matrix3d tfm = Matrix3d.MakeTranslation(100,50,0);
			parObj.DbEntity.Transform(tfm);
		}
	}
}
