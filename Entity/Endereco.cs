using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entity
{
    public class Endereco
    {
        public string Rua { get; set; }
        public string Bairro { get; set; }
        public string Cidade { get; set; }
        public string Estado { get; set; }
        public string Valor { get; set; }
        public int Periodo { get; set; }
        public int Vencimento { get; set; }
        public DateTime InicioLocacao { get; set; }
        public DateTime TerminoLocacao { get; set; }
    }
}
