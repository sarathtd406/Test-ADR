# Google Cloud Storage 

This documents defines the approach to select Google services based on specific organization requirements.

| Author | Status | Version | Date |
|:--|:--|:--|:--|
| Jane Doe | Accepted | 0.1 | 11-06-2023 |

<br>

## Context
----------

Google Cloud Storage (GCS) is a service for storing your objects in Google Cloud. An object is an immutable piece of data consisting of a file of any format. You store objects in containers called buckets. All buckets are associated with a project, and you can group your projects under an organization. 
GCS provides several storage classes that enable a consumer to choose the right balance of access frequency against cost. More frequently accessed storage has a higher storage cost but lower access charge, while less frequently accessed data has a lower storage charge and higher access charges.


## Decision
----------

- Need for an immutable compliant storage on the cloud to store regulatory related files and encrypt using customer managed encryption key (CMEK).
- The storage solution must be simple, scalable and can be accessed through APIs.

Usecases: 
- Media content storage and delivery. Geo-redundant storage with the highest level of availability.
- Backups and archives.
- durable storage for static websites.
- Integrated repository for analytics and ML.

GCS Object storage solution has a unique position with application transformation. GCS offers API level access on data storage for different storage requirements. GCS is tightly integrated with GCE, GKE, database services and git. 

It is recommend GCS Object storage as one of the storage solution options. 

## Consequences
----------

Pros or Capabilities:

- Key Functional Capabilities for Google Cloud Storage can be found [here](https://cloud.google.com/storage#section-9).

Cons or Service Limitations:

- The resource locations Organization Policy Service constraint controls the ability to create regional resources. When regional resource are created, testing should be carried out for new policy on non-production projects and folders.

- GCS Services explicitly to be used by Terraform Module for Prod/UAT and Development environment.

- Cloud Storage 5TB object size limit. Cloud Storage supports a maximum single-object size up 5 terabytes. If you have objects larger than 5TB, the object transfer fails for those objects for either Cloud Storage or Storage Transfer Service.

- Limit & Quotas about cloud Storage can be found [here](https://cloud.google.com/storage/quotas)

## Compliance
----------

- Encryption At Rest

  All Cloud Storage data are automatically stored in an encrypted state, but you can also provide your own encryption keys. For more information, see Cloud Storage Encryption.

- Encrypt your Cloud Storage data with Cloud KMS

- Turn on uniform bucket-level access and its org policy.

- Enforce the domain-restricted sharing constraint in an organization policy to prevent accidental public data sharing, or sharing beyond your organization.

- Audit your Cloud Storage data with Cloud Audit Logging.

- Secure your data with VPC Service Controls.

&nbsp;

#### ADR Changelog
----------

| Change Description | Version | Date |
|:--|:--|:--|
| Initial proposed version | 0.1 | 11-06-2023 |
