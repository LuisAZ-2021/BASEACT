using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Base
{
    public class RolItem
    {
        public int Id { get; set; }
        public string Rol { get; set; }

        public RolItem() { }

        public RolItem(int id, string nombre)
        {
            this.Id = id;
            this.Rol = nombre;
        }
    }
}
