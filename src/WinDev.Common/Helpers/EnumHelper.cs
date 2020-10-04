using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.ComponentModel;
using System.Reflection;
using System.Globalization;
using System.Data;
using System.Windows.Forms;

namespace Simergy.Common.Helpers
{
    public class EnumHelper
    {
        private const string DisplayName = "Value";
        private const string ValueName = "Key";

        public static string GetDescription(Enum value)
        {
            if (value == null)
            {
                throw new ArgumentNullException("value");
            }

            string description = value.ToString();
            FieldInfo fieldInfo = value.GetType().GetField(description);
            DescriptionAttribute[] attributes = (DescriptionAttribute[])fieldInfo.GetCustomAttributes(typeof(DescriptionAttribute), false);

            if (attributes != null && attributes.Length > 0)
            {
                description = attributes[0].Description;
            }
            return description;
        }

        public static T GetValueFromDescription<T>(string description)
        {
            if (!typeof(T).IsEnum)
            {
                throw new InvalidOperationException("The generic type T must be a valid enumeration.");
            }

            Array enumValues = Enum.GetValues(typeof(T));

            for (int i = 0; i < enumValues.Length; i++)
            {
                if (description.Equals(GetDescription((Enum)enumValues.GetValue(i))))
                {
                    return (T)(object)enumValues.GetValue(i);
                }
            }

            //foreach (Enum value in enumValues)
            //{
            //    if (description.Equals(GetDescription(value)))
            //    {
            //        return (T)(object)value;
            //    }
            //}

            throw new InvalidEnumArgumentException(string.Format(CultureInfo.CurrentCulture, "Could not find the description value '{0}' for the enumeration provided {1}.", description, typeof(T).Name));
        }

        public static int GetIndexOfValueFromCombobox(ComboBox comboBox, string value)
        {
            int index = 0;
            foreach (object item in comboBox.Items)
            {
                if (item is DataRowView)
                {
                    DataRowView dtv = item as DataRowView;

                    DataRow dr = dtv.Row;

                    if (dr.ItemArray[1].ToString() == value)
                        return index;

                    if (dr.ItemArray[0].ToString() == value)
                        return index;
                }
                else if (item is string)
                {
                    if (item.ToString() == value)
                        return index;
                }
                else if (item is Label)
                {
                    Label label = item as Label;
                    if (label.Text == value)
                        return index;
                }

                index++;
            }

            if (comboBox.Items.Count == 0)
                return -1;
            else
                return 0;
        }

        public static IList ToList<T>()
        {
            ArrayList list = new ArrayList();
            Array enumValues = Enum.GetValues(typeof(T));

            foreach (Enum value in enumValues)
            {
                list.Add(value.ToString());
            }

            return list;
        }

        public static void EnumToDropDown(Type enumType, ComboBox dropDown)
        {
            Array values = System.Enum.GetValues(enumType);

            foreach (int value in values)
            {
                string display = Enum.GetName(enumType, value);
                dropDown.Items.Add(display);

            }

        }        
    }
}
