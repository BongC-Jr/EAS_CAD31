using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
//using System.Windows.Forms;
using App = System.Windows.Forms.Application;

namespace EASI_CAD31
{
    public class GenaiMethods
    {
        static Document actDoc = Application.DocumentManager.MdiActiveDocument;
        static Database aCurDB = actDoc.Database;

        static CCData ccData = new CCData();
        static bool is_log = true;
        
        /**
         * Date added: 21 Jul 2025
         * Added by: Engr Bernardo Cabebe Jr.
         * Venue: 31H OOC, Bagumbayan, Quezon City
         */
        public static void DrawBWSection()
        {
            actDoc.Editor.WriteMessage($"\nDraw Builtup Section");
        }

        public static void AI_ConcBeamCapacity(List<object> loParam)
        {
            /**
             *AI_ConcBeamCapacity(new List<object>{beamW, beamH, nTensBar, nCompBar, barD, Mu, barDv, cFc, rfy,
             *                                     0      1      2         3         4     5   6      7    8
             *                    par1Index, iteratorValue, par2Index, targetValue})
             *                    9          10             11         12
             * parIndex is the index of a variable to be iterated
             * par2Index is the index of a variable to be compared to the target value
             * Example:
             *   concBeamCapacity(400,500,4,3,25,0.00,12,27.58,424,[2,1,5,870])
             *   
             *   In the example above, the nTensionBar will be incremented by 1 each iteration so that Mu at 
             *   at index 5 will achieve the targetValue.
             */
            double[] paramVariable = new double[] { Convert.ToDouble(loParam[0]), //beamW
                                                    Convert.ToDouble(loParam[1]), //beamH
                                                    Convert.ToDouble(loParam[2]), //nTensBar
                                                    Convert.ToDouble(loParam[3]), //nCompBar
                                                    Convert.ToDouble(loParam[4]), //barD
                                                    Convert.ToDouble(loParam[5])};//Mu

            //paramVariable[0] = 0 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + Convert.ToDouble(loParam[10]) : paramVariable[0];
            //paramVariable[1] = 1 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + Convert.ToDouble(loParam[10]) : paramVariable[1];
            //paramVariable[2] = 2 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + Convert.ToDouble(loParam[10]) : paramVariable[2];
            //paramVariable[3] = 3 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + Convert.ToDouble(loParam[10]) : paramVariable[3];
            //paramVariable[4] = 4 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + Convert.ToDouble(loParam[10]) : paramVariable[4];
            //paramVariable[5] = 5 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + Convert.ToDouble(loParam[10]) : paramVariable[5];

            string convo = "";
            // Validate the parameter index
            if (Convert.ToInt16(loParam[9]) < 0 || Convert.ToInt16(loParam[9]) > 5)
            {
                actDoc.Editor.WriteMessage($"Error: Index value = {Convert.ToInt16(loParam[9])}");
                convo = $"\nassistant: Error - there is something wrong with the index value which is {Convert.ToInt16(loParam[9])}\n";
                LogConversation(convo);

                actDoc.Editor.WriteMessage(convo);
                return;
            }
            if (Convert.ToInt16(loParam[11]) < 0 || Convert.ToInt16(loParam[11]) > 5)
            {
                actDoc.Editor.WriteMessage($"Error: Index value = {Convert.ToInt16(loParam[11])}");
                convo = $"\nassistant: Error - there is something wrong with the index value which is {Convert.ToInt16(loParam[11])}\n";
                LogConversation(convo);

                actDoc.Editor.WriteMessage(convo);
                return;
            }

            double cc = 40; //mm
            double ecu = 0.003;
            double Es = 200000; //MPa
            double barDv = Convert.ToDouble(loParam[6]);
            double cFc = Convert.ToDouble(loParam[7]);
            double rfy = Convert.ToDouble(loParam[8]);
            string cby = "Compresion bar yielded";

            double iterator = Convert.ToDouble(loParam[10]);//Increment value
            paramVariable[0] = 0 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + iterator : paramVariable[0];
            paramVariable[1] = 1 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + iterator : paramVariable[1];
            paramVariable[2] = 2 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + iterator : paramVariable[2];
            paramVariable[3] = 3 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + iterator : paramVariable[3];
            paramVariable[4] = 4 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + iterator : paramVariable[4];
            paramVariable[5] = 5 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + iterator : paramVariable[5];

            paramVariable[0] = 0 == Convert.ToInt16(loParam[11]) ? paramVariable[Convert.ToInt16(loParam[11])] : paramVariable[0];
            paramVariable[1] = 1 == Convert.ToInt16(loParam[11]) ? paramVariable[Convert.ToInt16(loParam[11])] : paramVariable[1];
            paramVariable[2] = 2 == Convert.ToInt16(loParam[11]) ? paramVariable[Convert.ToInt16(loParam[11])] : paramVariable[2];
            paramVariable[3] = 3 == Convert.ToInt16(loParam[11]) ? paramVariable[Convert.ToInt16(loParam[11])] : paramVariable[3];
            paramVariable[4] = 4 == Convert.ToInt16(loParam[11]) ? paramVariable[Convert.ToInt16(loParam[11])] : paramVariable[4];
            paramVariable[5] = 5 == Convert.ToInt16(loParam[11]) ? paramVariable[Convert.ToInt16(loParam[11])] : paramVariable[5];

            string logText = $"\npar2Index: {Convert.ToInt16(loParam[11])}";
            if (is_log)
            {
                WriteLog(logText);
            }
            

            bool isDiverging = false;
            int iterationCount = 0;
            double previousDifference = Math.Abs(Convert.ToDouble(loParam[12]) - paramVariable[Convert.ToInt16(loParam[11])]);
            logText = $"\npreviousDiff= {previousDifference}";
            if (is_log)
            {
                WriteLog(logText);
            }

            double bw = 0.0;
            double bh = 0.0;
            double nt = 0.0;
            double nc = 0.0;
            double bD = 0.0;
            double mc = 0.0;
            double dt = 0.0;
            double b1 = 0.0;
            double dc = 0.0;
            double Asc = 0.0;
            double rho = 0.0;

            do
            {
                // Check for divergence in the first two iterations
                double currentDifference = Math.Abs(Convert.ToDouble(loParam[12]) - paramVariable[Convert.ToInt16(loParam[11])]);
                
                if (iterationCount == 1)
                {
                    previousDifference = currentDifference;
                }
                else if (currentDifference > previousDifference)
                {
                    isDiverging = true;
                }
                previousDifference = currentDifference;
                
                logText = $"\ncurrentDiff= {currentDifference}";
                if (is_log)
                {
                    WriteLog(logText);
                }


                // Adjust the iterator value if diverging
                if (isDiverging)
                {
                    iterator = -iterator / 2;
                    isDiverging = false; // Reset the flag
                }
                logText = $"\niterator= {iterator}";
                if (is_log)
                {
                    WriteLog(logText);
                }

                // Update the variable to be incremented
                paramVariable[Convert.ToInt16(loParam[9])] += iterator;
                logText = $"\ndelta variable= {paramVariable[Convert.ToInt16(loParam[9])]}";
                if (is_log)
                {
                    WriteLog(logText);
                }

                // Check for acceptable tolerance
                if ((Math.Abs(Convert.ToDouble(loParam[12]) - paramVariable[Convert.ToInt16(loParam[11])]) <= 0.001) && iterationCount>0)
                {
                    break;
                }

                bw = 0 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + iterator : paramVariable[0];
                bh = 1 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + iterator : paramVariable[1];
                nt = 2 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + iterator : paramVariable[2];
                nc = 3 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + iterator : paramVariable[3];
                bD = 4 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + iterator : paramVariable[4];
                mc = 5 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + iterator : paramVariable[5];

                bw = 0 == Convert.ToInt16(loParam[11]) ? paramVariable[Convert.ToInt16(loParam[11])] : paramVariable[0];
                bh = 1 == Convert.ToInt16(loParam[11]) ? paramVariable[Convert.ToInt16(loParam[11])] : paramVariable[1];
                nt = 2 == Convert.ToInt16(loParam[11]) ? paramVariable[Convert.ToInt16(loParam[11])] : paramVariable[2];
                nc = 3 == Convert.ToInt16(loParam[11]) ? paramVariable[Convert.ToInt16(loParam[11])] : paramVariable[3];
                bD = 4 == Convert.ToInt16(loParam[11]) ? paramVariable[Convert.ToInt16(loParam[11])] : paramVariable[4];
                mc = 5 == Convert.ToInt16(loParam[11]) ? paramVariable[Convert.ToInt16(loParam[11])] : paramVariable[5];
                logText = $"\nbw={bw}\nbh={bh}\nnt={nt}\nnc={nc}";
                if (is_log)
                {
                    WriteLog(logText);
                }

                dc = cc + barDv + 0.5 * bD; //mm
                logText = $"\ndc={dc}";
                if (is_log)
                {
                    WriteLog(logText);
                }
                dt = bh - cc - barDv - 0.5 * bD; //mm
                logText = $"\nd={dt}";
                if (is_log)
                {
                    WriteLog(logText);
                }
                Asc = 0.25 * Math.PI * bD * bD * nc; //mm2
                logText = $"\nAsc={Asc}";
                if (is_log)
                {
                    WriteLog(logText);
                }
                double Ast = 0.25 * Math.PI * bD * bD * nt; //mm2
                rho = Ast / (bw * bh);
                logText = $"\nAst={Ast}\nrho={rho}";
                if (is_log)
                {
                    WriteLog(logText);
                }
                double Tf = Ast * rfy; //mm2*MPa = N
                logText = $"\nTf={Tf}";
                if (is_log)
                {
                    WriteLog(logText);
                }
                double rfs = rfy;//Assume fs=fy, MPa
                double Cs = (rfs - 0.85 * cFc) * Asc; //N
                logText = $"\nCs={Cs}";
                if (is_log)
                {
                    WriteLog(logText);
                }
                double ac = (Tf - Cs) / (0.85 * cFc * bw); //N/(MPa*mm) = mm
                b1 = Beta1(cFc); //-
                double x = ac / b1; //mm
                logText = $"\nx(init)={x}";
                if (is_log)
                {
                    WriteLog(logText);
                }
                //єsc=(x-dc)єcu/x
                double esc = (x - dc) * ecu / x; //mm
                logText = $"\nesc={esc}";
                if (is_log)
                {
                    WriteLog(logText);
                }
                //єy = fy / Es
                double ey = rfy / Es; //-
                if (esc < ey)
                {
                    cby = "Compresion bar not yielded";
                    //k1=	0.85f'cbβ1
                    double k1 = 0.85 * cFc * bw * b1; //N/mm
                    //k2 = (єcuEs - 0.85f'c)Asc-Asfy
                    double k2 = (ecu * Es - 0.85 * cFc) * Asc - Ast * rfy; //N
                    //k3=	dcєcuEsAsc
                    double k3 = dc * ecu * Es * Asc; //Nmm
                    //x=	[-k2+√(k22+4k1k3)]/2k1
                    x = (-k2 + Math.Sqrt(k2 * k2 + 4 * k1 * k3)) / (2 * k1); //mm
                    esc = (x - dc) * ecu / x; //-
                    rfs = esc * Es; //MPa
                }
                logText = $"\nx(final)={x}\n{cby}";
                if (is_log)
                {
                    WriteLog(logText);
                }
                ac = b1 * x; //mm
                //Cc=	0.85f'cba
                double Cc = 0.85 * cFc * bw * ac;
                logText = $"\nCc={Cc}";
                if (is_log)
                {
                    WriteLog(logText);
                }
                Cs = (rfs - 0.85 * cFc) * Asc;
                logText = $"\nCs={Cs}";
                if (is_log)
                {
                    WriteLog(logText);
                }

                //φMn = φ[Cc(d - 0.5a) + Cs(d - dc)]
                double phi = 0.9;
                mc = phi * (Cc * (dt - 0.5 * ac) + Cs * (dt - dc));
                logText = $"\nMc={mc}";
                if (is_log)
                {
                    WriteLog(logText);
                }

                paramVariable[0] = 0 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + iterator : bw;
                paramVariable[1] = 1 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + iterator : bh;
                paramVariable[2] = 2 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + iterator : nt;
                paramVariable[3] = 3 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + iterator : nc;
                paramVariable[4] = 4 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + iterator : bD;
                paramVariable[5] = 5 == Convert.ToInt16(loParam[9]) ? paramVariable[Convert.ToInt16(loParam[9])] + iterator : mc;

                paramVariable[0] = 0 == Convert.ToInt16(loParam[11]) ? paramVariable[Convert.ToInt16(loParam[11])] : bw;
                paramVariable[1] = 1 == Convert.ToInt16(loParam[11]) ? paramVariable[Convert.ToInt16(loParam[11])] : bh;
                paramVariable[2] = 2 == Convert.ToInt16(loParam[11]) ? paramVariable[Convert.ToInt16(loParam[11])] : nt;
                paramVariable[3] = 3 == Convert.ToInt16(loParam[11]) ? paramVariable[Convert.ToInt16(loParam[11])] : nc;
                paramVariable[4] = 4 == Convert.ToInt16(loParam[11]) ? paramVariable[Convert.ToInt16(loParam[11])] : bD;
                paramVariable[5] = 5 == Convert.ToInt16(loParam[11]) ? paramVariable[Convert.ToInt16(loParam[11])] : mc;

                logText = $"\n---{iterationCount}---------------\n\n";
                if (is_log)
                {
                    WriteLog(logText);
                }
                iterationCount++;
                
            } while ((Math.Abs(Convert.ToDouble(loParam[12]) - paramVariable[Convert.ToInt16(loParam[11])]) > 0.001) && iterationCount < 150);


            //xb=	єcud/(єcu+fy/Es)
            double xb = ecu * dt / (ecu + rfy / Es);
            double xmax = 0.75 * xb;
            double amax = b1 * xmax;
            //Cc=	0.85f'cb*amax
            double Ccmax = 0.85 * cFc * bw * amax;
            //єsc = (xmax - dc)єcu / xmax
            double escmax = (xmax - dc) * ecu / xmax;
            double rfsmax = rfy;
            if (escmax < (rfy / Es))
            {
                rfsmax = escmax * Es;
            }
            double Csmax = (rfsmax - 0.85 * cFc) * Asc;
            double Tmax = Ccmax + Csmax;
            double Asmax = Tmax / rfy; //mm2
            double rhomax = Asmax / (bw * bh);

            string ductile = "";
            if(rho > rhomax)
            {
                ductile = $"[rho={Math.Round(rho,5)}] <= [rhomax={Math.Round(rhomax,5)}], brittle failure";
            }
            else
            {
                ductile = $"[rho={Math.Round(rho,5)}] <= [rhomax={Math.Round(rhomax,5)}], ductile";
            }

            string mixText = "abcdefghijklmnopqrstuvwxyz0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_-";
            string shuffledText = ShuffleString(mixText);
            string first4chars = shuffledText.Substring(0, 4);
            convo = $"assistant: Result of calculation {first4chars}, the beam infos are width={Math.Round(paramVariable[0],2)} mm, " + 
                    $"depth={Math.Round(paramVariable[1],2)} mm, " +
                    $"tension bars={Math.Round(paramVariable[2],2)} pcs, compression bars={Math.Round(paramVariable[3],2)} pcs, " +
                    $"bar diameter={paramVariable[4]} mm, phiMn={Math.Round(paramVariable[5],4)} Nmm, f'c={cFc}, fy={rfy}, " +
                    $"stirrup dia.={barDv}, beta1={Math.Round(b1,3)}, {cby}, {ductile}\n";

            LogConversation(convo);
            actDoc.Editor.WriteMessage($"\n{convo}\n");
        }

        public static double Beta1(double fc)
        {
            double beta1 = (0.85 - 0.05 * (fc - 27) / 7);
            return Math.Min(Math.Max(beta1, 0.65), 0.85);
        }

        public static string ShuffleString(string input)
        {
            var chars = input.ToCharArray();
            var random = new Random();
            for (int i = chars.Length - 1; i > 0; i--)
            {
                int j = random.Next(i + 1);
                var temp = chars[i];
                chars[i] = chars[j];
                chars[j] = temp;
            }
            return new string(chars);
        }


        public void ExecuteMethod(string methodCallString)
        {
            try
            {
                // Extract method name and parameters
                string methodName = methodCallString.Substring(0, methodCallString.IndexOf('('));
                string parametersString = methodCallString.Substring(methodCallString.IndexOf('(') + 1, methodCallString.IndexOf(')') - methodCallString.IndexOf('(') - 1);

                // Get the method info
                MethodInfo methodInfo = typeof(GenaiMethods).GetMethod(methodName);

                if (methodInfo == null)
                {
                    throw new Exception($"Method {methodName} not found.");
                }

                // Parse parameters
                object[] parameters = ParseParameters(methodInfo, parametersString);

                // Invoke the method
                methodInfo.Invoke(null, parameters);
            }
            catch (TargetException ex)
            {
                actDoc.Editor.WriteMessage($"Target exception: {ex.Message}");
            }
            catch (TargetParameterCountException ex)
            {
                actDoc.Editor.WriteMessage($"Parameter count mismatch: {ex.Message}");
            }
            catch (ArgumentException ex)
            {
                actDoc.Editor.WriteMessage($"Argument exception: {ex.Message}");
            }
            catch (Exception ex)
            {
                actDoc.Editor.WriteMessage($"An error occurred: {ex.Message}");
            }
        }


        public static object[] ParseParameters(MethodInfo methodInfo, string parametersString)
        {
            try
            {
                ParameterInfo[] parameterInfos = methodInfo.GetParameters();
                object[] parameters = new object[parameterInfos.Length];

                // Assuming the parameters are in the format "new List<object>{400, 600, 6, 3, 25, 0.0, 10, 27.58, 414, 2, 6, 2, 6}"
                string listString = parametersString.Substring(parametersString.IndexOf('{') + 1, parametersString.IndexOf('}') - parametersString.IndexOf('{') - 1);
                string[] parameterStrings = listString.Split(',');

                if (parameterInfos.Length == 1 && parameterInfos[0].ParameterType == typeof(List<object>))
                {
                    List<object> list = new List<object>();
                    foreach (string parameterString in parameterStrings)
                    {
                        if (parameterString.Contains("."))
                        {
                            list.Add(Convert.ToDouble(parameterString.Trim()));
                        }
                        else if (parameterString == "true" || parameterString == "false")
                        {
                            list.Add(Convert.ToBoolean(parameterString.Trim()));
                        }
                        else
                        {
                            list.Add(Convert.ToInt32(parameterString.Trim()));
                        }
                    }
                    parameters[0] = list;
                }

                return parameters;
            }
            catch (Exception ex)
            {
                throw new Exception($"An error occurred while parsing parameters: {ex.Message}");
            }
        }


        public static void LogConversation(string conversationContent)
        {
            string filePath = Path.Combine(DataGlobal.convofilepath, "conversation.txt");
            string content = conversationContent;
            File.AppendAllText(filePath, content + Environment.NewLine);
        }

        public static void WriteLog(string logText)
        {
            string filePath = Path.Combine(DataGlobal.convofilepath, "logsH6p.txt");
            string content = logText;
            File.AppendAllText(filePath, content);
        }

    }
}
