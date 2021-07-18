using Entity;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Xceed.Words.NET;

namespace Documento.Model
{
    public class Documento
    {
        public string[] clausula =
            {
                "CLAUSULA PRIMEIRA:",
                "CLAUSULA SEGUNDA:",
                "CLAUSULA TERCEIRA:",
                "CLAUSULA QUARTA:",
                "CLAUSULA QUINTA:",
                "CLAUSULA SEXTA:",
                "CLAUSULA SÉTIMA:",
                "CLAUSULA OITIVA:",
                "CLAUSULA NOVA:",
                "CLAUSULA DÉCIMA:",
                "CLAUSULA DÉCIMA PRIMEIRA:",
                "CLAUSULA DÉCIMA SEGUNDA:",
                "CLAUSULA DÉCIMA TERCEIRA:",
                "CLAUSULA DÉCIMA QUARTA:",
                "CLAUSULA DÉCIMA QUINTA:",
                "CLAUSULA DÉCIMA SEXTA:",
                "CLAUSULA DÉCIMA SÉTIMA:",
                "CLAUSULA DÉCIMA OITAVA:",
                "TESTEMUNHAS",
            };

        public string[] dia =
        {
            "(um)",
            "(dois)",
            "(tres)",
            "(quatro)",
            "(cinco)",
            "(seis)",
            "(sete)",
            "(oito)",
            "(nove)",
            "(dez)",
            "(onze)",
            "(doze)",
            "(treze)",
            "(quatorze)",
            "(quinze)",
            "(dezesseis)",
            "(dezesete)",
            "(dezoito)",
            "(dezenove)",
            "(vinte)",
            "(vinte e um)",
            "(vinte e dois)",
            "(vinte e três)",
            "(vinte e quatro)",
            "(vinte e cinco)",
            "(vinte e seis)",
            "(vinte e sete)",
            "(vinte e oito)",
            "(vinte e nove)",
            "(trinta)",
            "(trinta e um)",
        };

        public string[] centena =
        {
            "CENTO",
            "DUZENTOS",
            "TREZENTOS",
            "QUATROCENTOS",
            "QUINHENTOS",
            "SEISCENTOS",
            "SETECENTOS",
            "OITOCENTOS",
            "NOVECENTOS",
            "MIL",
        };

        public string[] dezena =
        {
            "",
            " E DEZ",
            " E VINTE",
            " E TRINTA",
            " E QUARENTA",
            " E CINQUENTA",
            " E SESSENTA",
            " E SETENTA",
            " E OITENTA",
            " E NOVENTA",
        };

        public Documento(Pessoa locador, Pessoa locatario)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "DocX File|*.docx";
            dialog.Title = "Savar Contrato";
            dialog.ShowDialog();

            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string fileName = desktop + "\\" + locatario.CPF + "-" + locatario.Nome + ".docx";

            if (!String.IsNullOrEmpty(dialog.FileName)) fileName = dialog.FileName;

            fileName = fileName.Replace(" ", string.Empty);
            var docx = DocX.Create(fileName);

            InicioDocumento(docx, locador, locatario);

            CL_Primeira(docx, locatario);
            CL_Segunda(docx, locatario, locador);
            CL_Terceira(docx);
            CL_Quarta(docx);
            CL_Quinta(docx);
            CL_Sexta(docx);
            CL_Setima(docx);
            CL_Oitava(docx);
            CL_Nona(docx);
            CL_Decima(docx);
            CL_DecimaPrimeira(docx);
            CL_DecimaSegunda(docx);
            CL_DecimaTerceira(docx, locatario);
            CL_DecimaQuarta(docx);
            CL_DecimaQuinta(docx);
            CL_DecimaSexta(docx);
            CL_DecimaSetima(docx);
            CL_DecimaOitava(docx);

            FimDocumento(docx, locador);
            FimDocumento_Assinatura(docx, locador, locatario);
            FimDocumento_Testemunha(docx);

            docx.Save();
            docx.Dispose();

            if (File.Exists(fileName)) Process.Start("WINWORD.EXE", fileName);
        }

        private void InicioDocumento(DocX docx, Pessoa locador, Pessoa locatario)
        {
            string portarLocador = verificaSexoPessoa(locador);
            string portarLocatario = verificaSexoPessoa(locatario);

            //------------------------------------------------------

            string titulo = "CONTRATO DE LOCAÇÃO";
            string paragLocador = ", " + locador.Profissao + portarLocador + "do RG nº " + locador.RG + ", CPF " + locador.CPF +
                ", residente na " + locatario.Endereco.Rua + ", " + locatario.Endereco.Bairro + ", na cidade de " + locatario.Endereco.Cidade + ".";
            string paragLocatario = ", " + locatario.Profissao + portarLocatario + "do RG nº " + locatario.RG + ", CPF " + locatario.CPF +
                ", residente na " + locatario.Endereco.Rua + ", " + locatario.Endereco.Bairro + ", na cidade de " + locatario.Endereco.Cidade + ".";

            //-----------------------------------
            // INICIO CRIAR DOC
            //-----------------------------------

            docx.InsertParagraph(titulo).Font("Arial").FontSize(11).Bold().Alignment = Alignment.center;
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("LOCADOR (a): " + locador.Nome).Font("Arial").FontSize(11).Bold().Append(paragLocador).Font("Arial").FontSize(11)
                .SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("LOCATÁRIO (a): " + locatario.Nome).Font("Arial").FontSize(11).Bold().Append(paragLocatario).Font("Arial").FontSize(11)
                .SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void CL_Primeira(DocX docx, Pessoa locatario)
        {
            docx.InsertParagraph(clausula[0]).Font("Arial").FontSize(11).Bold().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("Finalidade da locação: Residencial").Font("Arial").FontSize(11).SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("Endereço do imóvel locado: ").Font("Arial").FontSize(11).Bold().Append(locatario.Endereco.Rua + ", " + locatario.Endereco.Bairro + ", " +
                locatario.Endereco.Cidade + "/" + locatario.Endereco.Estado).Font("Arial").FontSize(11).SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void CL_Segunda(DocX docx, Pessoa locatario, Pessoa locador)
        {
            string banco = locador.Pagamento.Banco;
            string pagamento;
            string preFrasePagamento;
            string posFrasePagamento;

            if (banco == "PESSOALMENTE")
            {
                preFrasePagamento = "ao Sr(a). ";
                pagamento =  locador.Nome;
                posFrasePagamento = ".";
            }
            else
            {
                preFrasePagamento = "em depósito ao Sr(a). ";
                pagamento =  locador.Nome + ", Banco: " + locador.Pagamento.Banco + ", Agencia: " + locador.Pagamento.Agencia + ", Conta Corrente: " +
                locador.Pagamento.Conta;
                posFrasePagamento = ", independente de aviso de recebimento.";
            }

            docx.InsertParagraph(clausula[1]).Font("Arial").FontSize(11).Bold().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("O valor mensal da locação será de R$ " + locatario.Endereco.Valor  + retornaValorExtenso(locatario.Endereco.Valor) + " mensais, cujo pagamento deverá ser efetuado " +
                "até o dia " + locatario.Endereco.Vencimento + " " + dia[locatario.Endereco.Vencimento - 1] + " de cada mês.").Font("Arial").FontSize(11).SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);

            docx.InsertParagraph("PARÁGRAFO PRIMEIRO: Prazo da locação: " + locatario.Endereco.Periodo + " " + dia[locatario.Endereco.Periodo - 1] + " meses, com início em " + 
                locatario.Endereco.InicioLocacao.ToString("dd/MM/yyyy") + " e término " + locatario.Endereco.TerminoLocacao.ToString("dd/MM/yyyy") + ". " + 
                "Os aluguéis serão ajustados anualmente pelo IGPM. Após 12 meses de efetivada a locação convindo aos locatários, " +
                "poderão rescindir a locação com o expresso aviso de desocupação com 03(três) meses de antecedência.").Font("Arial").FontSize(11).IndentationBefore = 1.0f;
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);

            docx.InsertParagraph("PARAGRAFO SEGUNDO: A presente locação reger - se á pelo Código Civil " +
                "Brasileiro e pela Lei do Inquilinato, sendo seu valor locatício corrigido de acordo com a normas legais. Fica desde já convencionado entre as partes que caso haja " +
                "alteração na legislação, os locatários desde já declaram aceitar a se submeterem às novas regras permitidas pela lei;").Font("Arial").FontSize(11)
                .IndentationBefore = 1.0f;
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);

            docx.InsertParagraph("PARAGRAFO TERCEIRO: O aluguel ora estipulado deverá ser pago até o dia estabelecido de cada mês vencido, diretamente ").Font("Arial").FontSize(11)
                .Append(preFrasePagamento).Font("Arial").FontSize(11).Append(pagamento).Font("Arial").FontSize(11).Bold().Append(posFrasePagamento)
                .Font("Arial").FontSize(11).IndentationBefore = 1.0f;
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void CL_Terceira(DocX docx)
        {
            docx.InsertParagraph(clausula[2]).Font("Arial").FontSize(11).Bold().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("O não pagamento do aluguel no dia do vencimento acarretará ao locatário o pagamento de multa de 10% (dez por cento) sobre o valor do aluguel, " +
                "mais juros e correção diariamente atualizada, até o seu efetivo pagamento.").Font("Arial").FontSize(11).SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void CL_Quarta(DocX docx)
        {
            docx.InsertParagraph(clausula[3]).Font("Arial").FontSize(11).Bold().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("Além do aluguel estipulado na Cláusula Segunda, ficará a cargo do LOCATÁRIO, despesas com melhorias e taxas que recaem ou que venham a recair sobre o " +
                "imóvel, sendo estas já existentes ou que venha a ser criadas, inclusive as majorações, bem como as despesas de água e luz, cujos pagamentos deverão ser feitos " +
                "nos locais apropriados para o seu estabelecimento, e devidamente comprovados ao LOCADOR.").Font("Arial").FontSize(11).SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void CL_Quinta(DocX docx)
        {
            docx.InsertParagraph(clausula[4]).Font("Arial").FontSize(11).Bold().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("O LOCATÁRIO declara ter vistoriado o imóvel objeto da presente locação, aceitando-o nas condições em que se encontra, ficando ainda, " +
                "obrigados a: ").Font("Arial").FontSize(11).SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);

            var list = docx.AddList("Mantê-lo no mais perfeito estado de higiene, conservação e limpeza, com acessórios de iluminação, aparelhos sanitários, pintura, ralos, telhados, " +
                "portas e janelas, não sendo permitido a colocação de PREGOS E BUCHAS NAS PAREDES, para assim restituir quando finda a locação;", 1, ListItemType.Bulleted);
            docx.AddListItem(list, "Satisfazer por suas expensas, sem direito à indenização ou retenção, toda e qualquer exigência dos poderes públicos a que derem causa; mormente pelas modificações " +
                "e adaptações que realizarem;", 1);
            docx.AddListItem(list, "Fazer por sua conta as reparações de estragos que não provenham do uso normal do imóvel ou a que tiver dado causa;", 1);
            docx.AddListItem(list, "A usá-lo única e exclusivamente para o fim residencial.", 1);


            docx.InsertList(list, new Xceed.Words.NET.Font("Arial"), 11);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void CL_Sexta(DocX docx)
        {
            docx.InsertParagraph(clausula[5]).Font("Arial").FontSize(11).Bold().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("As benfeitorias introduzidas no imóvel pelo LOCATÁRIO, ainda que necessárias, desde logo se incorporarão no imóvel, sem que tenha direito à " +
                "indenização ou retenção quando finda ou rescindida a locação.").Font("Arial").FontSize(11).SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void CL_Setima(DocX docx)
        {
            docx.InsertParagraph(clausula[6]).Font("Arial").FontSize(11).Bold().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("O LOCATÁRIO faculta, desde já, ao LOCADOR, vistoriar o imóvel quando assim entender conveniente, mediante estipulação de dia e hora.")
                .Font("Arial").FontSize(11).SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void CL_Oitava(DocX docx)
        {
            docx.InsertParagraph(clausula[7]).Font("Arial").FontSize(11).Bold().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("O LOCATÁRIO se compromete a fazer chegar às mãos do LOCADOR os avisos de comunicações que digam respeito ao imóvel locado e cujo encargo não seja de " +
                "sua responsabilidade, em razão da presente avença, sob pena de responderem por perdas e danos.").Font("Arial").FontSize(11).SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void CL_Nona(DocX docx)
        {
            docx.InsertParagraph(clausula[8]).Font("Arial").FontSize(11).Bold().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("Em caso de desapropriação de imóvel, a locação será considerada rescindida, não cabendo ao LOCADOR ressarcir qualquer prejuízo eventualmente " +
                "sofrido pelo LOCATÁRIO.").Font("Arial").FontSize(11).SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void CL_Decima(DocX docx)
        {
            docx.InsertParagraph(clausula[9]).Font("Arial").FontSize(11).Bold().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("Em caso de incêndio ou acidente que exija a reconstrução do imóvel, fica rescindido este contrato, independente da multa contratual, " +
                "com responsabilidade do LOCATÁRIO se os fatos forem imputados.").Font("Arial").FontSize(11).SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void CL_DecimaPrimeira(DocX docx)
        {
            docx.InsertParagraph(clausula[10]).Font("Arial").FontSize(11).Bold().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("Nenhuma intimação do poder público será motivo para a rescisão do presente contrato, salvo, procedendo à vistoria judicial que vislumbre a " +
                "imprestabilidade do imóvel para o fim que se destina.").Font("Arial").FontSize(11).SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void CL_DecimaSegunda(DocX docx)
        {
            docx.InsertParagraph(clausula[11]).Font("Arial").FontSize(11).Bold().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("A parte que infringir qualquer cláusula deste contrato pagará multa de 03 (três) aluguéis vigentes na ocasião, com faculdade para a parte " +
                "inocente considerar rescindida a locação, independentemente de qualquer notificação judicial ou extrajudicial. A multa será sempre paga integralmente, seja " +
                "qual for o tempo decorrido do presente contrato.").Font("Arial").FontSize(11).SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void CL_DecimaTerceira(DocX docx, Pessoa locatario)
        {
            docx.InsertParagraph(clausula[12]).Font("Arial").FontSize(11).Bold().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("Com expressa renúncia de qualquer outro, por mais privilegiado que seja, fica eleito o foro da Comarca de " + locatario.Endereco.Cidade +
                " /" + locatario.Endereco.Estado + " para todas as ações ou medidas judiciais cabíveis em razão deste contrato. A parte vencida em demanda pagará, " +
                "além da multa contratual, as custas processuais e honorários advocatícios da parte vencedora, fixados pelo juiz da causa.").Font("Arial").FontSize(11)
                .SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void CL_DecimaQuarta(DocX docx)
        {
            docx.InsertParagraph(clausula[13]).Font("Arial").FontSize(11).Bold().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("O recebimento de aluguéis e demais encargos da locação, fora do prazo ou por valor inferior ao pactuado por este contrato, representará mera " +
                "tolerância do LOCADOR, não constituindo, em hipótese alguma, novação, renovação ou alteração das cláusulas contratuais.").Font("Arial").FontSize(11)
                .SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void CL_DecimaQuinta(DocX docx)
        {
            docx.InsertParagraph(clausula[14]).Font("Arial").FontSize(11).Bold().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("Após o termo final do presente contrato, os aluguéis continuarão sofrendo reajustes com os índices legais, até a renovação contratual ou a entrega " +
                "das chaves. E para o caso do LOCATÁRIO tiver sido notificado para a desocupação no prazo final e não o fizer, incorrerá na penalidade prevista no artigo 575 do " +
                "Código Civil Brasileiro.").Font("Arial").FontSize(11).SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void CL_DecimaSexta(DocX docx)
        {
            docx.InsertParagraph(clausula[15]).Font("Arial").FontSize(11).Bold().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("Não é permitida a transferência deste contrato, no todo ou em parte do imóvel, nem a sublocação, cessão ou empréstimo, sem prévio consentimento, " +
                "por escrito, do LOCADOR, devendo, no caso da permissão, estarem todos obrigados a respeitarem as cláusulas constantes neste contrato.").Font("Arial").FontSize(11)
                .SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void CL_DecimaSetima(DocX docx)
        {
            docx.InsertParagraph(clausula[16]).Font("Arial").FontSize(11).Bold().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("Transmutando-se o presente contrato para o prazo indeterminado, por força da Lei, e não havendo interesse das partes na sua continuação, " +
                "deverá denunciar, por escrito, e com prazo mínimo de 30(trinta) dias, dando ciência inequívoca do desinteresse do seu prosseguimento.").Font("Arial").FontSize(11)
                .SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void CL_DecimaOitava(DocX docx)
        {
            docx.InsertParagraph(clausula[17]).Font("Arial").FontSize(11).Bold().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("O LOCATÁRIO se obriga a cumprir todas as obrigações constantes no Artigo 23, da Lei do Inquilino, no qual declara ter pleno conhecimento para " +
                "restituir o imóvel quando findo ou rescindido o presente contrato.").Font("Arial").FontSize(11).SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph("E, por estarem assim justos e contratados, assim este instrumento em(2) duas vias de igual teor e valor, na presença das testemunhas abaixo, " +
                "obrigando - se as partes, por si, seus herdeiros e sucessores.").Font("Arial").FontSize(11).SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void FimDocumento(DocX docx, Pessoa locador)
        {
            docx.InsertParagraph(locador.Endereco.Cidade + ", " + data()).Font("Arial").FontSize(11)
                .SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void FimDocumento_Assinatura(DocX docx, Pessoa locador, Pessoa locatario)
        {
            Table t = docx.AddTable(3, 2);
            var whiteBorder = new Border(Xceed.Words.NET.BorderStyle.Tcbs_none, 0, 0, Color.White);

            t.Rows[0].Cells[0].Paragraphs.First().Append("____________________________________").Font("Arial").FontSize(11).Alignment = Alignment.center;
            t.Rows[0].Cells[1].Paragraphs.First().Append("____________________________________").Font("Arial").FontSize(11).Alignment = Alignment.center;

            t.Rows[1].Cells[0].Paragraphs.First().Append(locador.Nome).Bold().Font("Arial").FontSize(11).Alignment = Alignment.center;
            t.Rows[1].Cells[1].Paragraphs.First().Append(locatario.Nome).Bold().Font("Arial").FontSize(11).Alignment = Alignment.center;

            t.Rows[2].Cells[0].Paragraphs.First().Append("LOCADOR").Bold().Font("Arial").FontSize(11).Alignment = Alignment.center;
            t.Rows[2].Cells[1].Paragraphs.First().Append("LOCATÁRIO").Bold().Font("Arial").FontSize(11).Alignment = Alignment.center;

            //----------------------------------------------------
            // BORDER STYLE

            t.SetBorder(TableBorderType.Bottom, whiteBorder);
            t.SetBorder(TableBorderType.Left, whiteBorder);
            t.SetBorder(TableBorderType.Right, whiteBorder);
            t.SetBorder(TableBorderType.Top, whiteBorder);
            t.SetBorder(TableBorderType.InsideH, whiteBorder);
            t.SetBorder(TableBorderType.InsideV, whiteBorder);

            //----------------------------------------------------

            docx.InsertTable(t);

            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);
        }

        private void FimDocumento_Testemunha(DocX docx)
        {
            docx.InsertParagraph(clausula[18]).Font("Arial").FontSize(11).Bold().SetLineSpacing(LineSpacingType.Line, 1.5f);
            docx.InsertParagraph().SetLineSpacing(LineSpacingType.Line, 1.5f);

            Table u = docx.AddTable(3, 2);
            var whiteBorder = new Border(Xceed.Words.NET.BorderStyle.Tcbs_none, 0, 0, Color.White);

            u.Rows[0].Cells[0].Paragraphs.First().Append("____________________________________").Font("Arial").FontSize(11).Alignment = Alignment.center;
            u.Rows[0].Cells[1].Paragraphs.First().Append("____________________________________").Font("Arial").FontSize(11).Alignment = Alignment.center;

            u.Rows[1].Cells[0].Paragraphs.First().Append("NOME: ").Bold().Font("Arial").FontSize(11).SetLineSpacing(LineSpacingType.Line, 1.5f);
            u.Rows[1].Cells[1].Paragraphs.First().Append("NOME: ").Bold().Font("Arial").FontSize(11).SetLineSpacing(LineSpacingType.Line, 1.5f);

            u.Rows[2].Cells[0].Paragraphs.First().Append("CPF: ").Bold().Font("Arial").FontSize(11).SetLineSpacing(LineSpacingType.Line, 1.5f);
            u.Rows[2].Cells[1].Paragraphs.First().Append("CPF: ").Bold().Font("Arial").FontSize(11).SetLineSpacing(LineSpacingType.Line, 1.5f);

            //----------------------------------------------------
            // BORDER STYLE

            u.SetBorder(TableBorderType.Bottom, whiteBorder);
            u.SetBorder(TableBorderType.Left, whiteBorder);
            u.SetBorder(TableBorderType.Right, whiteBorder);
            u.SetBorder(TableBorderType.Top, whiteBorder);
            u.SetBorder(TableBorderType.InsideH, whiteBorder);
            u.SetBorder(TableBorderType.InsideV, whiteBorder);

            //----------------------------------------------------

            docx.InsertTable(u);
        }

        public string verificaSexoPessoa(Pessoa pessoa)
        {
            if (pessoa.IsMasculino) return ", portador ";
            else return ", portadora ";
        }

        public string retornaValorExtenso(string valor)
        {
            string valorExtenso = "(";

            char[] digitos = valor.Replace(".00", "").ToCharArray();

            if(digitos.Length == 3)
            {
                valorExtenso += centena[Convert.ToInt16(digitos[0].ToString()) - 1];
                valorExtenso += dezena[Convert.ToInt16(digitos[1].ToString())] + " REAIS)";
            }

            return valorExtenso;
        }

        public string data()
        {
            CultureInfo culture = new CultureInfo("pt-BR");

            string mes = DateTime.Now.ToString("MMMM", culture);
            mes = char.ToUpper(mes[0]) + mes.Substring(1);

            string dateNow = DateTime.Now.Day.ToString() + " de " + mes + " de " + DateTime.Now.Year.ToString();

            return dateNow;
        }
    }
}
