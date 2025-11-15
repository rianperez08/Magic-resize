import { S3Client, PutObjectCommand } from "@aws-sdk/client-s3";

const hasS3 = !!process.env.S3_BUCKET;

export async function uploadFromUrlToS3(url, key, contentType) {
  if (!hasS3) return url; // fallback: return Canva temporary URL
  const r = await fetch(url);
  if (!r.ok) throw new Error("download_failed");
  const buf = Buffer.from(await r.arrayBuffer());
  const s3 = new S3Client({ region: process.env.AWS_REGION || "ap-southeast-1" });
  await s3.send(new PutObjectCommand({
    Bucket: process.env.S3_BUCKET,
    Key: key, Body: buf, ContentType: contentType, ACL: "private"
  }));
  return `s3://${process.env.S3_BUCKET}/${key}`;
}