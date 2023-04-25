using System;
using System.Collections.Generic;
using System.IO;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;

namespace Carga_Asignacion_Coopeuch_Castigo
{
    class Program
    {
        public static List<Asignacion_Coopeuch_Castigo> Lista_Asignacion_Coopeuch_Castigos = new List<Asignacion_Coopeuch_Castigo>();
        static void Main(string[] args)
        {

            /*-------------------------------------------------------------------------*/
            /*                             RUTA DEL ARCHIVO                            */
            /*-------------------------------------------------------------------------*/
            DirectoryInfo di = new DirectoryInfo(@"C:\coopeuch\asignacion\");
            FileInfo[] files = di.GetFiles("*.xlsx");

            foreach (FileInfo file in files)
            {
                Console.WriteLine("entre al forach");
                if (file.Name.Contains("Base_Castigo"))
                {
                    Console.WriteLine("entre al if");
                    String Ruta_Archivo = @"C:\coopeuch\asignacion\" + file.Name;
                    Console.WriteLine("Lectura del Archivo: " + file.Name.ToString());
                    /*-------------------------------------------------------------------------*/
                    /*                    llamado al metodo leer archivo                       */
                    /*-------------------------------------------------------------------------*/
                    Leer_Excel(Ruta_Archivo);
                }
            }

            int cantidad = 0;
            foreach (var i in Lista_Asignacion_Coopeuch_Castigos)
            {
                Console.WriteLine("Insertando Registro: " + cantidad.ToString());
                /*-------------------------------------------------------------------------*/
                /*                      CARGA A LA BASE DE DATOS                           */
                /*-------------------------------------------------------------------------*/
                //string connstring = @"Data Source=192.168.0.77; Initial Catalog=EJFDES; Persist Security Info=True; User ID=sa; Password=Desa2019;";
                string connstring = @"Data Source=192.168.0.5; Initial Catalog=EJFDES; Persist Security Info=True; User ID=sa; Password=w2003ejf103;";
                using (SqlConnection con = new SqlConnection(connstring))
                {
                    con.Open();
                    string commandString = @"INSERT INTO [dbo].[Coopeuch_Base_Castigos]
                        ([Operación],[Oficina],[Rut socio],[Dv],[Nombre socio],[Comuna Socio]
                        ,[Fecha Castigo],[Grupo Credito],[Nombre Producto],[Acciones]
                        ,[Saldo IBS (cierre mes anterior)],[Dirección],[Complemento Direccion]
                        ,[Nombre Comuna],[Celular],[Telefono Primario],[Telefono Secundario]
                        ,[Direccion Email],[INTERES MORA],[TASA],[SCRIPT])
                        VALUES(@OPERACION,@OFICINA,@RUT_SOCIO,@DV,@NOMBRE_SOCIO,
                        @COMUNA_SOCIO,@FECHA_CASTIGO,@GRUPO_CREDITO,@NOMBRE_PRODUCTO,@ACCIONES,
                        @SALDO_IBS,@DIRECCION,@COMPLEMENTO_DIRECCION,@NOMBRE_COMUNA,@CELULAR,
                        @TELEFONO_PRIMARIO,@TELEFONO_SECUNDARIO,@DIRECCION_EMAIL,
                        @INTERES_MORA,@TASA,NULL)";

                    SqlCommand cmd = new SqlCommand(commandString, con);
                    cmd.Parameters.AddWithValue("@OPERACION", i.OPERACION);
                    //cmd.Parameters.AddWithValue("@BASE", i.BASE);
                    //cmd.Parameters.AddWithValue("@AVENIMIENTO", i.AVENIMIENTO);
                    //cmd.Parameters.AddWithValue("@OBSERVACION", i.OBSERVACION);
                    //cmd.Parameters.AddWithValue("@VIGENCIA", i.VIGENCIA);
                    //cmd.Parameters.AddWithValue("@CONTRALOR", i.CONTRALOR);
                    //cmd.Parameters.AddWithValue("@GESTOR_MES_CURSO", i.GESTOR_MES_CURSO);
                    //cmd.Parameters.AddWithValue("@SUSPENSON_LEY_DE_QUIEBRA", i.SUSPENSION_LEY_DE_QUIEBRA);
                    //cmd.Parameters.AddWithValue("@COD_CONVENIO", i.COD_CONVENIO);
                    cmd.Parameters.AddWithValue("@OFICINA", i.OFICINA);
                    //cmd.Parameters.AddWithValue("@GERENCIA_REGIONAL", i.GERENCIA_REGIONAL);
                    cmd.Parameters.AddWithValue("@RUT_SOCIO", i.RUT_SOCIO);
                    cmd.Parameters.AddWithValue("@DV", i.DV);
                    cmd.Parameters.AddWithValue("@NOMBRE_SOCIO", i.NOMBRE_SOCIO);
                    cmd.Parameters.AddWithValue("@COMUNA_SOCIO", i.COMUNA_SOCIO);
                    cmd.Parameters.AddWithValue("@FECHA_CASTIGO", i.FECHA_CASTIGO);
                    //cmd.Parameters.AddWithValue("@ID_CASTIGO", i.ID_CASTIGO);
                    //cmd.Parameters.AddWithValue("@ANIO_CASTIGO", i.AÑO_CASTIGO);
                    cmd.Parameters.AddWithValue("@GRUPO_CREDITO", i.GRUPO_CREDITO);
                    cmd.Parameters.AddWithValue("@NOMBRE_PRODUCTO", i.NOMBRE_PRODUCTO);
                    //cmd.Parameters.AddWithValue("@CUOTAS_PACTADAS", i.CUOTAS_PACTADAS);
                    //cmd.Parameters.AddWithValue("@CUOTAS_PAGADAS", i.CUOTAS_PAGADAS);
                    //cmd.Parameters.AddWithValue("@VALOR_CUOTA", i.VALOR_CUOTA);
                    cmd.Parameters.AddWithValue("@ACCIONES", i.ACCIONES);
                    //cmd.Parameters.AddWithValue("@GASTO_JUDICIAL", i.GASTO_JUDICIAL);
                    //cmd.Parameters.AddWithValue("@MTO_CASTIGO_FA_CASTIGO", i.MTO_CASTIGO_FA_CASTIGO);
                    //cmd.Parameters.AddWithValue("@MTO_CASTIGO_CTA_5400", i.MTO_CASTIGO_CTA_5400);
                    //cmd.Parameters.AddWithValue("@ABONOS_AL_CIERRE_MES_ANTERIOR", i.ABONOS_AL_CIERRE_MES_ANTERIOR);
                    cmd.Parameters.AddWithValue("@SALDO_IBS", i.SALDO_IBS_CIERRE_MES_ANTERIOR);
                    //cmd.Parameters.AddWithValue("@MTO_RECUPERADO_CONTABLE", i.MTO_RECUPERADO_CONTABLE);
                    //cmd.Parameters.AddWithValue("@FECHA_RECUPERO", i.FECHA_RECUPERO);
                    //cmd.Parameters.AddWithValue("@FECHA_ULTIMO_ABONO", i.FECHA_ULTIMO_ABONO);
                    //cmd.Parameters.AddWithValue("@MONTO_ULTIMO_ABONO", i.MONTO_ULTIMO_ABONO);
                    //cmd.Parameters.AddWithValue("@RECUPERO_CON_INTERES", i.RECUPERO_CON_INTERES);
                    //cmd.Parameters.AddWithValue("@REASIGNACION", i.REASIGNACION);
                    //cmd.Parameters.AddWithValue("@SALDO_URANO", i.SALDO_URANO);
                    //cmd.Parameters.AddWithValue("@TIPO_RECUPERO", i.TIPO_RECUPERO);
                    cmd.Parameters.AddWithValue("@DIRECCION", i.DIRECCION);
                    cmd.Parameters.AddWithValue("@COMPLEMENTO_DIRECCION", i.COMPLEMENTO_DIRECCION);
                    cmd.Parameters.AddWithValue("@NOMBRE_COMUNA", i.NOMBRE_COMUNA);
                    cmd.Parameters.AddWithValue("@CELULAR", i.CELULAR);
                    cmd.Parameters.AddWithValue("@TELEFONO_PRIMARIO", i.TELEFONO_PRIMARIO);
                    cmd.Parameters.AddWithValue("@TELEFONO_SECUNDARIO", i.TELEFONO_SECUNDARIO);
                    cmd.Parameters.AddWithValue("@DIRECCION_EMAIL", i.DIRECCION_EMAIL);
                    //cmd.Parameters.AddWithValue("@SEXO", i.SEXO);
                    //cmd.Parameters.AddWithValue("@TRAMO_EDAD", i.TRAMO_EDAD);
                    //cmd.Parameters.AddWithValue("@TIPO_DEUDA", i.TIPO_DEUDA);
                    //cmd.Parameters.AddWithValue("@PAGA_PAGA", i.PAGA_PAGA);
                    //cmd.Parameters.AddWithValue("@3_PAGOS_ULTIMOS_6_MESES", i.PAGOS_3_ULTIMOS_6_MESES);
                    //cmd.Parameters.AddWithValue("@PAGO_MES_ANTERIOR", i.PAGO_MES_ANTERIOR);
                    //cmd.Parameters.AddWithValue("@FLUJO", i.FLUJO);
                    //cmd.Parameters.AddWithValue("@2_PAGOS_VIGENTE", i.PAGOS_2_VIGENTE);
                    //cmd.Parameters.AddWithValue("@PUNTAJE_MODELO", i.PUNTAJE_MODELO);
                    cmd.Parameters.AddWithValue("@INTERES_MORA", i.INTERES_MORA);
                    cmd.Parameters.AddWithValue("@TASA", i.TASA);

                    cmd.ExecuteNonQuery();
                    con.Close();
                    cantidad++;

                }
            }
            //Console.ReadKey();

            Lista_Asignacion_Coopeuch_Castigos = null;
            GC.Collect();
        }
        private static void Leer_Excel(string ruta)
        {
            int contador = 1;
            string Campo = "";
            DateTime Fecha_Actual = DateTime.Now;
            /*--------------------------------------------------------------------------------*/
            /*                              LECTURA DE ARCHIVO                                */
            /*--------------------------------------------------------------------------------*/
            try
            {
                using (var stream = File.Open(ruta, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration()
                    {
                        FallbackEncoding = Encoding.UTF8,
                        LeaveOpen = false,
                        AnalyzeInitialCsvRows = 0,
                    }))
                    {
                        do
                        {
                            //recorro el excel
                            while (reader.Read())
                            {
                                //omito cabecera
                                if (contador > 1)
                                {
                                    Asignacion_Coopeuch_Castigo Input = new Asignacion_Coopeuch_Castigo();

                                    Campo = "FECHA_CARGA";
                                    Input.FECHA_CARGA = Fecha_Actual;
                                    Campo = "OPERACION";
                                    Input.OPERACION = reader.GetValue(0) != null ? reader.GetValue(0).ToString() : "";
                                    Campo = "BASE";
                                    Input.BASE = "";/*reader.GetValue(2) != null ? reader.GetValue(3).ToString() : "";*/
                                    Campo = "AVENIMIENTO";
                                    Input.AVENIMIENTO = "";/*reader.GetValue(3) != null ? reader.GetValue(2).ToString() : "";*/
                                    Campo = "OBSERVACION";
                                    Input.OBSERVACION = "";/*reader.GetValue(4) != null ? reader.GetValue(4).ToString() : "";*/
                                    Campo = "VIGENCIA";
                                    Input.VIGENCIA = "";/*reader.GetValue(5) != null ? reader.GetValue(5).ToString() : "";*/
                                    Campo = "CONTRALOR";
                                    Input.CONTRALOR = "";/*reader.GetValue(6) != null ? reader.GetValue(6).ToString() : "";*/
                                    Campo = "GESTOR_MES_CURSO";
                                    Input.GESTOR_MES_CURSO = "";/*reader.GetValue(7) != null ? reader.GetValue(7).ToString() : "";*/
                                    Campo = "SUSPENSION_LEY_DE_QUIEBRA";
                                    Input.SUSPENSION_LEY_DE_QUIEBRA = "";/*reader.GetValue(8) != null ? reader.GetValue(8).ToString() : "";*/
                                    Campo = "COD_CONVENIO";
                                    Input.COD_CONVENIO = "";/*reader.GetValue(9) != null ? reader.GetValue(9).ToString() : "";*/
                                    Campo = "OFICINA";
                                    Input.OFICINA = reader.GetValue(10) != null ? reader.GetValue(10).ToString() : "";
                                    Campo = "GERENCIA_REGIONAL";
                                    Input.GERENCIA_REGIONAL = "";/*reader.GetValue(11) != null ? reader.GetValue(11).ToString() : "";*/
                                    Campo = "RUT_SOCIO";
                                    Input.RUT_SOCIO = reader.GetValue(12) != null ? int.Parse(reader.GetValue(12).ToString()) : 0;
                                    Campo = "DV";
                                    Input.DV = reader.GetValue(13) != null ? reader.GetValue(13).ToString() : "";
                                    Campo = "NOMBRE_SOCIO";
                                    Input.NOMBRE_SOCIO = reader.GetValue(14) != null ? reader.GetValue(14).ToString() : "";
                                    Campo = "COMUNA_SOCIO";
                                    Input.COMUNA_SOCIO = reader.GetValue(15) != null ? reader.GetValue(15).ToString() : "";
                                    Campo = "FECHA_CASTIGO";
                                    Input.FECHA_CASTIGO = reader.GetValue(16) != null ? DateTime.Parse(reader.GetValue(16).ToString()) : DateTime.Parse("1900-01-01");
                                    Campo = "ID_CASTIGO";
                                    Input.ID_CASTIGO = 0;/*reader.GetValue(17) != null ? float.Parse(reader.GetValue(17).ToString()) : 0;*/
                                    Campo = "AÑO_CASTIGO";
                                    Input.ANIO_CASTIGO = 0;/*reader.GetValue(18) != null ? float.Parse(reader.GetValue(18).ToString()) : 0;*/
                                    Campo = "GRUPO_CREDITO";
                                    Input.GRUPO_CREDITO = reader.GetValue(19) != null ? reader.GetValue(19).ToString() : "";
                                    Campo = "NOMBRE_PRODUCTO";
                                    Input.NOMBRE_PRODUCTO = reader.GetValue(20) != null ? reader.GetValue(20).ToString() : "";
                                    Campo = "CUOTAS_PACTADAS";
                                    Input.CUOTAS_PACTADAS = 0;/*reader.GetValue(21) != null ? float.Parse(reader.GetValue(21).ToString()) : 0;*/
                                    Campo = "CUOTAS_PAGADAS";
                                    Input.CUOTAS_PAGADAS = 0;/*reader.GetValue(22) != null ? float.Parse(reader.GetValue(22).ToString()) : 0;*/
                                    Campo = "VALOR_CUOTA";
                                    Input.VALOR_CUOTA = 0;/*reader.GetValue(23) != null ? float.Parse(reader.GetValue(23).ToString()) : 0;*/
                                    Campo = "ACCIONES";
                                    Input.ACCIONES = reader.GetValue(24) != null ? float.Parse(reader.GetValue(24).ToString()) : 0;
                                    Campo = "GASTO_JUDICIAL";
                                    Input.GASTO_JUDICIAL = 0;/*reader.GetValue(25) != null ? float.Parse(reader.GetValue(25).ToString()) : 0;*/
                                    Campo = "MTO_CASTIGO_FA_CASTIGO";
                                    Input.MTO_CASTIGO_FA_CASTIGO = 0;/*reader.GetValue(26) != null ? float.Parse(reader.GetValue(26).ToString()) : 0;*/
                                    Campo = "MTO_CASTIGO_CTA_5400";
                                    Input.MTO_CASTIGO_CTA_5400 = 0;/* reader.GetValue(27) != null ? float.Parse(reader.GetValue(27).ToString()) : 0;*/
                                    Campo = "ABONOS_AL_CIERRE_MES_ANTERIOR";
                                    Input.ABONOS_AL_CIERRE_MES_ANTERIOR = 0;/* reader.GetValue(28) != null ? float.Parse(reader.GetValue(28).ToString()) : 0;*/
                                    Campo = "SALDO_IBS_CIERRE_MES_ANTERIOR";
                                    Input.SALDO_IBS_CIERRE_MES_ANTERIOR = reader.GetValue(29) != null ? float.Parse(reader.GetValue(29).ToString()) : 0;
                                    Campo = "MTO_RECUPERADO_CONTABLE";
                                    Input.MTO_RECUPERADO_CONTABLE = 0;/* reader.GetValue(30) != null ? float.Parse(reader.GetValue(30).ToString()) : 0;*/
                                    Campo = "FECHA_RECUPERO";
                                    Input.FECHA_RECUPERO = "";/*reader.GetValue(31) != null ? reader.GetValue(31).ToString() : "";*/
                                    Campo = "FECHA_ULTIMO_ABONO";
                                    Input.FECHA_ULTIMO_ABONO = DateTime.Parse("1900-01-01");/*reader.GetValue(32) != null ? DateTime.Parse(reader.GetValue(32).ToString()) : DateTime.Parse("1900-01-01");*/
                                    Campo = "MONTO_ULTIMO_ABONO";
                                    Input.MONTO_ULTIMO_ABONO = 0;/* reader.GetValue(33) != null ? float.Parse(reader.GetValue(33).ToString()) : 0;*/
                                    Campo = "RECUPERO_CON_INTERES";
                                    Input.RECUPERO_CON_INTERES = 0;/* reader.GetValue(34) != null ? float.Parse(reader.GetValue(34).ToString()) : 0;*/
                                    Campo = "REASIGNACION";
                                    Input.REASIGNACION = ""; /* reader.GetValue(35) != null ? reader.GetValue(35).ToString() : "";*/
                                    Campo = "SALDO_URANO";
                                    Input.SALDO_URANO = 0;/* reader.GetValue(36) != null ? float.Parse(reader.GetValue(36).ToString()) : 0;*/
                                    Campo = "TIPO_RECUPERO";
                                    Input.TIPO_RECUPERO = "";/*reader.GetValue(37) != null ? reader.GetValue(37).ToString() : "";*/
                                    Campo = "DIRECCION";
                                    Input.DIRECCION = reader.GetValue(38) != null ? reader.GetValue(38).ToString() : "";
                                    Campo = "COMPLEMENTO_DIRECCION";
                                    Input.COMPLEMENTO_DIRECCION = reader.GetValue(39) != null ? reader.GetValue(39).ToString() : "";
                                    Campo = "NOMBRE_COMUNA";
                                    Input.NOMBRE_COMUNA = reader.GetValue(40) != null ? reader.GetValue(40).ToString() : "";
                                    Campo = "CELULAR";
                                    Input.CELULAR = reader.GetValue(41) != null ? reader.GetValue(41).ToString() : "";
                                    Campo = "TELEFONO_PRIMARIO";
                                    Input.TELEFONO_PRIMARIO = reader.GetValue(42) != null ? reader.GetValue(42).ToString() : "";
                                    Campo = "TELEFONO_SECUNDARIO";
                                    Input.TELEFONO_SECUNDARIO = reader.GetValue(43) != null ? reader.GetValue(43).ToString() : "";
                                    Campo = "DIRECCION_EMAIL";
                                    Input.DIRECCION_EMAIL = reader.GetValue(44) != null ? reader.GetValue(44).ToString() : "";
                                    Campo = "SEXO";
                                    Input.SEXO = "";/*reader.GetValue(45) != null ? reader.GetValue(45).ToString() : "";*/
                                    Campo = "TRAMO_EDAD";
                                    Input.TRAMO_EDAD = "";/*reader.GetValue(46) != null ? reader.GetValue(46).ToString() : "";*/
                                    Campo = "TIPO_DEUDA";
                                    Input.TIPO_DEUDA = "";/*reader.GetValue(47) != null ? reader.GetValue(47).ToString() : "";*/
                                    Campo = "PAGA_PAGA";
                                    Input.PAGA_PAGA = "";/*reader.GetValue(48) != null ? reader.GetValue(48).ToString() : "";*/
                                    Campo = "PAGOS_3_ULTIMOS_6_MESES";
                                    Input.PAGOS_3_ULTIMOS_6_MESES = "";/*reader.GetValue(49) != null ? reader.GetValue(49).ToString() : "";*/
                                    Campo = "PAGO_MES_ANTERIOR";
                                    Input.PAGO_MES_ANTERIOR = "";/*reader.GetValue(50) != null ? reader.GetValue(50).ToString() : "";*/
                                    Campo = "FLUJO";
                                    Input.FLUJO = "";/*reader.GetValue(51) != null ? reader.GetValue(51).ToString() : "";*/
                                    Campo = "PAGOS_2_VIGENTE";
                                    Input.PAGOS_2_VIGENTE = "";/*reader.GetValue(52) != null ? reader.GetValue(52).ToString() : "";*/
                                    Campo = "PUNTAJE_MODELO";
                                    Input.PUNTAJE_MODELO = 0;/*reader.GetValue(53) != null ? float.Parse(reader.GetValue(53).ToString()) : 0;*/
                                    Campo = "INTERES_MORA";
                                    Input.INTERES_MORA = reader.GetValue(57) != null ? int.Parse(reader.GetValue(57).ToString()) : 0;
                                    Campo = "TASA";
                                    Input.TASA = reader.GetValue(58) != null ? float.Parse(reader.GetValue(58).ToString()) : 0;

                                    Lista_Asignacion_Coopeuch_Castigos.Add(Input);
                                }
                                contador++;
                            }
                        } while (reader.NextResult());
                    }
                }
            }
            catch (Exception e)
            {
                string mensaje = "Campo: " + Campo + " contador: " + contador + " error: " + e;
                Console.WriteLine(mensaje);
            }
        }

        public class Asignacion_Coopeuch_Castigo
        {
            public DateTime FECHA_CARGA { get; set; }
            public String OPERACION { get; set; }
            public String BASE { get; set; }
            public String AVENIMIENTO { get; set; }
            public String OBSERVACION { get; set; }
            public String VIGENCIA { get; set; }
            public String CONTRALOR { get; set; }
            public String GESTOR_MES_CURSO { get; set; }
            public String SUSPENSION_LEY_DE_QUIEBRA { get; set; }
            public String COD_CONVENIO { get; set; }
            public String OFICINA { get; set; }
            public String GERENCIA_REGIONAL { get; set; }
            public int RUT_SOCIO { get; set; }
            public String DV { get; set; }
            public String NOMBRE_SOCIO { get; set; }
            public String COMUNA_SOCIO { get; set; }
            public DateTime FECHA_CASTIGO { get; set; }
            public float ID_CASTIGO { get; set; }
            public float ANIO_CASTIGO { get; set; }
            public String GRUPO_CREDITO { get; set; }
            public String NOMBRE_PRODUCTO { get; set; }
            public float CUOTAS_PACTADAS { get; set; }
            public float CUOTAS_PAGADAS { get; set; }
            public float VALOR_CUOTA { get; set; }
            public float ACCIONES { get; set; }
            public float GASTO_JUDICIAL { get; set; }
            public float MTO_CASTIGO_FA_CASTIGO { get; set; }
            public float MTO_CASTIGO_CTA_5400 { get; set; }
            public float ABONOS_AL_CIERRE_MES_ANTERIOR { get; set; }
            public float SALDO_IBS_CIERRE_MES_ANTERIOR { get; set; }
            public float MTO_RECUPERADO_CONTABLE { get; set; }
            public String FECHA_RECUPERO { get; set; }
            public DateTime FECHA_ULTIMO_ABONO { get; set; }
            public float MONTO_ULTIMO_ABONO { get; set; }
            public float RECUPERO_CON_INTERES { get; set; }
            public String REASIGNACION { get; set; }
            public float SALDO_URANO { get; set; }
            public String TIPO_RECUPERO { get; set; }
            public String DIRECCION { get; set; }
            public String COMPLEMENTO_DIRECCION { get; set; }
            public String NOMBRE_COMUNA { get; set; }
            public String CELULAR { get; set; }
            public String TELEFONO_PRIMARIO { get; set; }
            public String TELEFONO_SECUNDARIO { get; set; }
            public String DIRECCION_EMAIL { get; set; }
            public String SEXO { get; set; }
            public String TRAMO_EDAD { get; set; }
            public String TIPO_DEUDA { get; set; }
            public String PAGA_PAGA { get; set; }
            public String PAGOS_3_ULTIMOS_6_MESES { get; set; }
            public String PAGO_MES_ANTERIOR { get; set; }
            public String FLUJO { get; set; }
            public String PAGOS_2_VIGENTE { get; set; }
            public float PUNTAJE_MODELO { get; set; }
            public int INTERES_MORA { get; set; }
            public float TASA { get; set; }

        }

    }
}
