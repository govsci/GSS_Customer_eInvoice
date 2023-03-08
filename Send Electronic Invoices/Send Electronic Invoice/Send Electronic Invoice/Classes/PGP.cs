using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Org.BouncyCastle.Security;    // PGP Key Library
using Org.BouncyCastle.Bcpg.OpenPgp;
using Org.BouncyCastle.Utilities.IO;
using Org.BouncyCastle.Bcpg;

namespace Send_Electronic_Invoice.Classes
{
    class PGP
    {
        private string outputFilePath;

        public string GetFilePath()
        {
            return outputFilePath;
        }

        private static PgpPublicKey ReadPublicKey(Stream inputStream)
        {

            inputStream = PgpUtilities.GetDecoderStream(inputStream);
            PgpPublicKeyRingBundle pgpPub = new PgpPublicKeyRingBundle(inputStream);

            foreach (PgpPublicKeyRing kRing in pgpPub.GetKeyRings())
            {
                foreach (PgpPublicKey k in kRing.GetPublicKeys())
                {
                    if (k.IsEncryptionKey)
                    {
                        return k;
                    }
                }
            }

            throw new ArgumentException("Can't find encryption key in key ring.");

        }

        /**
        * Search a secret key ring collection for a secret key corresponding to
        * keyId if it exists.
        *
        * @param pgpSec a secret key ring collection.
        * @param keyId keyId we want.
        * @param pass passphrase to decrypt secret key with.
        * @return
        */

        private static void EncryptFile(Stream outputStream, string fileName, PgpPublicKey encKey, bool armor, bool withIntegrityCheck)
        {

            if (armor)
            {

                outputStream = new ArmoredOutputStream(outputStream);

            }

            try
            {

                MemoryStream bOut = new MemoryStream();

                PgpCompressedDataGenerator comData = new PgpCompressedDataGenerator(CompressionAlgorithmTag.Zip);

                PgpUtilities.WriteFileToLiteralData(

                comData.Open(bOut),

                PgpLiteralData.Binary,

                new FileInfo(fileName));

                comData.Close();

                PgpEncryptedDataGenerator cPk = new PgpEncryptedDataGenerator(

                SymmetricKeyAlgorithmTag.Cast5, withIntegrityCheck, new SecureRandom());

                cPk.AddMethod(encKey);

                byte[] bytes = bOut.ToArray();

                Stream cOut = cPk.Open(outputStream, bytes.Length);

                cOut.Write(bytes, 0, bytes.Length);

                cOut.Close();

                if (armor)
                {

                    outputStream.Close();

                }

            }

            catch (PgpException e)
            {

                Console.Error.WriteLine(e);

                Exception underlyingException = e.InnerException;

                if (underlyingException != null)
                {

                    Console.Error.WriteLine(underlyingException.Message);

                    Console.Error.WriteLine(underlyingException.StackTrace);

                }

            }

        }


        public void Encrypt(string file, string pathToSave, string publicKeyFile, string fileName, string fileExtention)
        {
            string year = DateTime.Now.ToString("yyyy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string path = $"{pathToSave}\\{year}\\{month}\\{day}\\";

            outputFilePath = $"{path}{fileName.Replace(".csv", "")}{fileExtention}";
            Stream keyIn;

            keyIn = File.OpenRead(publicKeyFile);

            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            using (Stream fos = File.Create(outputFilePath))
                EncryptFile(fos, file, ReadPublicKey(keyIn), true, true);

            keyIn.Close();
        }
    }
}
