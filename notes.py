# use inpspect.getmembers() for use with new packages (very helpful)

#for getting the time of signature
tmp_time_t = trust_verification_result.GetTimeOfTrustVerification()

            trust_verification_time_enum = trust_verification_result.GetTimeOfTrustVerificationEnum()

            if trust_verification_time_enum is VerificationOptions.e_current:
                print("Trust verification attempted with respect to current time (as epoch time): " + str(tmp_time_t))
            elif trust_verification_time_enum is VerificationOptions.e_signing:
                print("Trust verification attempted with respect to signing time (as epoch time): " + str(tmp_time_t))
            elif trust_verification_time_enum is VerificationOptions.e_timestamp:
                print("Trust verification attempted with respect to secure embedded timestamp (as epoch time): " + str(
                    tmp_time_t))
            else:
                assert False, "unrecognized time enum value"

#for getting any changes or annotations to the file
changes = result.GetDisallowedChanges()
        for it2 in changes:
            print("\tDisallowed change: %s, objnum: %lu" % (it2.GetTypeAsString(), it2.GetObjNum()))

#honestly i dont know what function this performs
#first line in for loop trusts the certificate,
for i in in_public_key_file_path:
    opts.AddTrustedCertificate(sigs_path + i)
    result = curr.Verify(opts)
    if result.GetVerificationStatus():
        print(i)
        print("Signature verified, objnum: %lu" % curr.GetSDFObj().GetObjNum())
        break
    else:
        print(i)
        print("Signature verification failed, objnum: %lu" % curr.GetSDFObj().GetObjNum())
        verification_status = False

# Open an existing PDF
doc = PDFDoc(in_docpath)

# Choose a security level to use, and change any verification options you wish to change
opts = VerificationOptions(VerificationOptions.e_compatibility_and_archiving)

# Add trust root to store of trusted certificates contained in VerificationOptions.
opts.AddTrustedCertificate(in_public_key_file_path)

result = doc.VerifySignedDigitalSignatures(opts)



print('inserting name')
                                subject[sig_names] = subject_dn[1].GetStringValue()
                                print('incrementing name')

# useless bollocks
def verify_all_and_print(in_docpath, in_public_key_file_path):
    print("==========")

    print("\t\t\t=============" + in_docpath + "=========")
    doc = PDFDoc(in_docpath)

    subject = []
    sig_count = 0
    sig_names = 0

    digsig_fitr = doc.GetDigitalSignatureFieldIterator()
    while digsig_fitr.HasNext():

        for i in in_public_key_file_path:
            opts = VerificationOptions(VerificationOptions.e_compatibility_and_archiving)
            trusted_cert = (r".\sigs\\" + i)
            opts.AddTrustedCertificate(trusted_cert)

            curr = digsig_fitr.Current()
            result = curr.Verify(opts)

            if result.GetVerificationStatus():
                print("signature can be verified")
                if result.HasTrustVerificationResult():
                    trust_result = result.GetTrustVerificationResult()
                    print(trust_result.GetResultString)
                    if trust_result.WasSuccessful:
                        print("Certificate Verified")
                        sig_count = sig_count + 1
                        if not trust_result.GetCertPath():
                            print("Could not print certificate path.")
                        else:
                            cert_path = trust_result.GetCertPath()
                            for j in range(len(cert_path)):
                                print("\tCertificate:")
                                full_cert = cert_path[j]
                                print("\t\tIssuer names:")
                                issuer_dn = full_cert.GetIssuerField().GetAllAttributesAndValues()
                                for k in range(len(issuer_dn)):
                                    print("\t\t\t" + issuer_dn[k].GetStringValue())

                                print("\t\tSubject names:")
                                subject_dn = full_cert.GetSubjectField().GetAllAttributesAndValues()
                                subject = full_cert.GetSubjectField().GetAllAttributesAndValues()
                                for q in subject:
                                    q = q.GetStringValue()
                                sig_names = sig_names + 1
                                for s in subject_dn:
                                    print("\t\t\t" + s.GetStringValue())

                                print("\t\tExtensions:")
                                for x in full_cert.GetExtensions():
                                    print("\t\t\t" + x.ToString())
                    else:
                        print("certificate not verified")
                        print(trust_result.GetResultString())
                else:
                    print("result has no verification status")
            else:
                print("cannot verify signiture")
            print("===============================Finished====================")
        digsig_fitr.Next()
        print("=================Next Field===========================================")

# returns all the paths of PDFs in the directory specified by paths
    def get_paths(self, path):
        pdf_list = []
        for (dirpath, dirnames, filenames) in os.walk(path):
            pdf_list += [os.path.join(dirpath, file) for file in filenames if file.endswith('.pdf')]
        return pdf_list

# removes duplicate items from both lists
    def remove_blacklisted_paths(self, blacklist, pdf_list):
        for element in blacklist:
            if element in pdf_list:
                pdf_list.remove(element)
        return pdf_list