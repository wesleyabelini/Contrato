using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entity
{
    public class Pessoa
    {
        public string Nome { get; set; }
        public string RG { get; set; }
        public string CPF { get; set; }
        public Endereco Endereco { get; set; }
        public bool IsMasculino { get; set; }
        public string Profissao { get; set; }
        public Pagamento Pagamento { get; set; }
    }
}
