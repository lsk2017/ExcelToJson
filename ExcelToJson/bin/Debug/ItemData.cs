

// auto generated file

using System;
using System.Collections.Generic;

namespace TableData
{
	public readonly struct ItemData
	{
		public readonly int id;
		public readonly string name;
		public readonly int dp;
		public readonly int mr;
		public readonly ItemType type;

		public ItemData(JObject json)
		{
			name = json["name"].ToString();

			id = int.Parse(json["id"].ToString());
			dp = int.Parse(json["dp"].ToString());
			mr = int.Parse(json["mr"].ToString());

			type = (ItemType)json["type"];
		}
	}
}